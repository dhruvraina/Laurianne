%fgwnh
close all; clear all;

%Instructions for use:
% 1. Make sure the 'CurrentFolder' (underneath 'Window' and 'Help' in the main toolbar) is pointing to 'F:\Dhruv\laurianne'
% 2. Change 'MainDirectory' and 'OutputDirectory' if you need to
% 3. Press the 'Save' Button at the top of this window (the little floppy disk icon) if you make any changes!
% 4. In the 'Command Window' (the window directly below this one) type 'laurianne.m' and press enter.


%This is all you need to edit! Make sure all the images are 5012x5012!
%Don't set the output directory the same as the input directory
MainDirectory = 'C:\Users\ncbs\Downloads\French girl who needs help\';
OutputDirectory = 'C:\Users\ncbs\Desktop\fgwnh_outputs\';


%% Main Program
dir_lst = dir(MainDirectory);
for ctr2 = 3:length(dir_lst)
    imdir = [MainDirectory dir_lst(ctr2).name];
    im = bfopen(imdir);
    series1 = im{1,1};
    
    chan1 = series1{1};
    chan2 = series1{2};
    chan3 = series1{3};
    
    chan1r = imresize(chan1, [512 512]);
    chan1r = mat2gray(imtophat(chan1r,strel('disk', 30)), [0 255]);
    chan2r = mat2gray(imresize(chan2, [512 512]), [0 255]);
    chan3r = mat2gray(imresize(chan3, [512 512]), [0 255]);
    
    tt_flat = reshape(chan1r, [512*512, 1]);
    [idx, c, sumcent] = kmeans(tt_flat,3, 'emptyaction', 'drop');
    if max(isnan(c))
        [idx, ~, sumcent] = kmeans(tt_flat,3, 'emptyaction', 'drop');
        z=1;
        msgbox(['Please check the following image and re-run it if segmentation looks poor!: ' dir_lst(ctr2).name(1:end-4) '  The segmentation sometimes fails because of the inherent nature of the algorithm (k-means)'] )
    end
    reshape_im = reshape(idx, [512 512]);
    filt(1) = mode(tt_flat(idx==1));
    filt(2) = mode(tt_flat(idx==2));
    filt(3) = mode(tt_flat(idx==3));
    bckg = min(filt);
    throw = find(filt==bckg);
    bb=chan1r;
    bb(reshape_im==throw)=0;
    
    bb = im2bw(bb, 0.0001);
    qq = bb.*chan1r;
    qq_bw = im2bw(qq, 1.4*graythresh(qq));
    qq_clear = bwareaopen(qq_bw,2);
    qq_clear = imfill(qq_clear, 'holes');
    qq_clear = imclearborder(qq_clear);
    imop = imerode(qq_clear, strel('disk', 2));
      
    [lbl num] = bwlabel(imop);
    lbl = imdilate(lbl, strel('disk', 3));
        
    bwfinal = zeros(512,512);
    bwfinal(lbl>0)=1;
    
    rp_ch1 = regionprops(lbl, chan1r,  'Area', 'MeanIntensity' );
    rp_ch2 = regionprops(lbl, chan2r, 'MeanIntensity' );
    rp_ch3 = regionprops(lbl, chan3r, 'MeanIntensity' );
    
    for ctr1 = 1:length(rp_ch1)
        resultmat(ctr1,1) = rp_ch1(ctr1).Area;
        resultmat(ctr1,2) = rp_ch1(ctr1).MeanIntensity*255;
        resultmat(ctr1,3) = rp_ch2(ctr1).MeanIntensity*255;
        resultmat(ctr1,4) = rp_ch3(ctr1).MeanIntensity*255;
    end
    
    chan2im = chan2r.*bwfinal;
    chan1im = chan1r.*bwfinal;
    chan3im = chan3r.*bwfinal;
    
   %% Last bit: Saving Images, Writing to Excel

    figure, imagesc(chan1im)
    set(gcf, 'visible', 'off');
    hgexport(gcf, [OutputDirectory dir_lst(ctr2).name(1:end-4) '_Channel1' '.jpg'], hgexport('factorystyle'), 'Format', 'jpeg');
    close gcf;
    
    figure, imagesc(chan2im)
    set(gcf, 'visible', 'off');
    hgexport(gcf, [OutputDirectory dir_lst(ctr2).name(1:end-4) '_Channel2' '.jpg'], hgexport('factorystyle'), 'Format', 'jpeg');
    close gcf;
    
    figure, imagesc(chan3im)
    set(gcf, 'visible', 'off');
    hgexport(gcf, [OutputDirectory dir_lst(ctr2).name(1:end-4) '_Channel3' '.jpg'], hgexport('factorystyle'), 'Format', 'jpeg');
    close gcf;
    
    figure, imagesc(chan1r)
    set(gcf, 'visible', 'off');
    hgexport(gcf, [OutputDirectory dir_lst(ctr2).name(1:end-4) '_Channel1_OriginalImage' '.jpg'], hgexport('factorystyle'), 'Format', 'jpeg');
    close gcf;
    
    outvec = {'Relative Cell Area' 'Channel1 Mean Intensity' 'Channel2 Mean Intensity' 'Channel3 Mean Intensity' };
    for ctr3 = 1:length(resultmat)
        mm = ctr3+1;%offset header
        outvec{mm,1} = resultmat(ctr3,1);
        outvec{mm,2} = resultmat(ctr3,2);
        outvec{mm,3} = resultmat(ctr3,3);
        outvec{mm,4} = resultmat(ctr3,4);
    end
    try
        xlswrite([OutputDirectory dir_lst(ctr2).name(1:end-4) '.xlsx'], outvec);
    catch
        errordlg('The output excel file may be in use or the disk may be write protected. Please close all MS Excel windows or select an alternate output folder and try again.')
    end
    clearvars -except MainDirectory OutputDirectory dir_lst ctr2
end
msgbox('Done!')

