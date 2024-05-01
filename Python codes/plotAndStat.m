% We are plotting graphs and stats for the given data for 2-5 and 6-9 [4 channels at a time] So you have to run the code twice
% 1. Uncomment line 10 and comment line 11
% 2. Uncomment line 41 and comment line 40
% 3. The plots will be saved in matlab folder . Refer to the flowchart for the folder structure

% Define output folder path
outputFolder = 'OutputResults_2-5';
%outputFOlder =  'OutputResults_6-9'; 

% Create output folder if it doesn't exist
if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

% Iterate over folders from OutputExcel/25/ to OutputExcel/36/
for folderNum = 13:24
    % Define data folder path for the current iteration
    dataFolder = fullfile('OutputExcel', num2str(folderNum));
    
    % Iterate over DataFor1 and DataFor2 subfolders
    for dataSubfolder = {'DataFor1', 'DataFor2'}
        subfolderName = dataSubfolder{1};
        subfolder = fullfile(outputFolder, sprintf('DataFor%d_%s_Results', folderNum, subfolderName));
        
        % Create subfolder for each data subfolder if it doesn't exist
        if ~exist(subfolder, 'dir')
            mkdir(subfolder);
        end
        
        % Iterate over Excel files in the current data subfolder
        for fileIdx = 0:6
            filename = fullfile(dataFolder, subfolderName, sprintf('%d.xlsx', fileIdx));
            try
                % Read SEMG signals from Excel file into a table
                T = readtable(filename);
                t = T.(1); % time vector in seconds
                signals = table2array(T(:, 2:5)); % assuming the SEMG signals are in columns 2 to 5
                %signals = table2array(T(:, 6:9)); % assuming the SEMG signals are in columns 6 to 9
                
                % Design a 4th order Butterworth band-stop filter
                fs = 1000; % sampling frequency in Hz
                f1 = 48; % lower frequency cutoff in Hz
                f2 = 52; % upper frequency cutoff in Hz
                Wn = [f1/(fs/2), f2/(fs/2)];
                [b,a] = butter(4, Wn, 'stop');
                
                % Initialize arrays to store results
                max_amp = zeros(4,1); % maximum amplitude of amplitude spectrum
                peak_freq = zeros(4,1); % frequency with maximum PSD
                mean_amp = zeros(4,1); % mean of the amplitude spectrum
                mean_time = zeros(4,1); % mean in time domain
                skewness_time = zeros(4,1); % skewness in time domain
                kurtosis_time = zeros(4,1); % kurtosis in time domain
                
                % Apply the filter to each signal and plot in subplots
                figure;
                for i = 1:4
                    signal = signals(:,i);
                    filtered_signal = filter(b, a, signal);
                    
                    % Plot the filtered signal and its frequency spectrum
                    subplot(4,2,(i-1)*2+1);
                    plot(t, filtered_signal);
                    xlabel('Time (ms)');
                    ylabel('Amplitude (mV)');
                    title(sprintf('Filtered Signal %d', i));
                    
                    nfft = 2^nextpow2(length(filtered_signal));
                    f = linspace(0, fs/2, nfft/2+1);
                    filtered_signal_fft = fft(filtered_signal, nfft);
                    filtered_signal_fft_mag = abs(filtered_signal_fft(1:nfft/2+1));
                    
                    % Find maximum amplitude of amplitude spectrum and peak frequency
                    [max_amp(i), idx] = max(filtered_signal_fft_mag(2:end));
                    peak_freq(i) = f(idx+1);
                    mean_amp(i) = mean(filtered_signal_fft_mag);
                    
                    % Plot frequency spectrum of filtered signal
                    subplot(4,2,i*2);
                    plot(f, filtered_signal_fft_mag);
                    xlim([0, 100]);
                    xlabel('Frequency (Hz)');
                    ylabel('Magnitude');
                    title(sprintf('Frequency Spectrum of Filtered Signal %d', i));
                    
                    % Calculate mean, skewness, and kurtosis in time domain
                    mean_time(i) = mean(filtered_signal);
                    skewness_time(i) = skewness(filtered_signal);
                    kurtosis_time(i) = kurtosis(filtered_signal);
                end
                
                % Save statistical measures in a text file
                resultsFilename = fullfile(subfolder, sprintf('Results_%d.txt', fileIdx));
                fid = fopen(resultsFilename, 'w');
                fprintf(fid, "Maximum amplitude of amplitude spectrum: \n");
                fprintf(fid, "%f\n", max_amp);
                fprintf(fid, "Peak frequency (excluding 0 frequency component): \n");
                fprintf(fid, "%f\n", peak_freq);
                fprintf(fid, "Mean of amplitude spectrum: \n");
                fprintf(fid, "%f\n", mean_amp);
                fprintf(fid, "Mean in time domain: \n");
                fprintf(fid, "%f\n", mean_time);
                fprintf(fid, "Skewness in time domain: \n");
                fprintf(fid, "%f\n", skewness_time);
                fprintf(fid, "Kurtosis in time domain: \n");
                fprintf(fid, "%f\n", kurtosis_time);
                fclose(fid);
                
                % Save the plot
                saveas(gcf, fullfile(subfolder, sprintf('Plot_%d.png', fileIdx)));
                close(gcf);
            catch
                disp(['Error processing file: ' filename]);
            end
        end
    end
end
