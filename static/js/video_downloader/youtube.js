(function() {
    function init() {
    const analyzeBtn = document.getElementById('analyze-btn');
    const urlInput = document.getElementById('video-url');
    if (!analyzeBtn || !urlInput) return;
    const loadingState = document.getElementById('loading-state');
    const errorAlert = document.getElementById('error-alert');
    const errorMessage = document.getElementById('error-message');
    const resultsSection = document.getElementById('results-section');
    
    // Metadata elements
    const videoThumbnail = document.getElementById('video-thumbnail');
    const videoDuration = document.getElementById('video-duration');
    const videoTitle = document.getElementById('video-title');
    
    // Tables
    const videoFormatsContainer = document.getElementById('video-formats-container');
    const videoFormatsBody = document.getElementById('video-formats-body');
    const audioFormatsContainer = document.getElementById('audio-formats-container');
    const audioFormatsBody = document.getElementById('audio-formats-body');
    
    function formatBytes(bytes, decimals = 2) {
        if (!+bytes) return 'Unknown';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }
    
    function formatDuration(seconds) {
        if (!seconds || seconds <= 0) return 'Unavailable';
        const h = Math.floor(seconds / 3600);
        const m = Math.floor((seconds % 3600) / 60);
        const s = Math.floor(seconds % 60);
        if (h > 0) {
            return h + ':' + m.toString().padStart(2, '0') + ':' + s.toString().padStart(2, '0');
        }
        return m + ':' + s.toString().padStart(2, '0');
    }
    
    urlInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            analyzeBtn.click();
        }
    });
    
    urlInput.addEventListener('input', function() {
        if (!errorAlert.classList.contains('hidden')) {
            errorAlert.classList.add('hidden');
        }
    });
    
    analyzeBtn.addEventListener('click', async function() {
        const url = urlInput.value.trim();
        if (!url) {
            urlInput.focus();
            return;
        }
        
        // Reset state
        errorAlert.classList.add('hidden');
        resultsSection.classList.add('hidden');
        loadingState.classList.remove('hidden');
        analyzeBtn.disabled = true;
        analyzeBtn.innerHTML = '<i data-lucide="loader-2" class="w-5 h-5 animate-spin"></i><span>Analyzing</span>';
        if(window.lucide) window.lucide.createIcons();
        
        try {
            const analyzeEndpoint = window.API_ENDPOINTS?.analyze || '/video-downloader/api/analyze/';
            const response = await fetch(analyzeEndpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ url: url })
            });
            
            const data = await response.json();
            
            if (!response.ok) {
                throw new Error(data.message || data.error || 'Failed to analyze video');
            }
            
            // Populate metadata
            videoTitle.textContent = data.title;
            videoDuration.textContent = 'Duration: ' + formatDuration(data.duration);
            if (data.thumbnail) {
                videoThumbnail.src = data.thumbnail;
            } else {
                videoThumbnail.src = window.STATIC_URLS?.placeholder || '';
            }
            
            // Clear tables
            videoFormatsBody.innerHTML = '';
            audioFormatsBody.innerHTML = '';
            
            let hasVideo = false;
            let hasAudio = false;
            
            const sortedFormats = data.formats.sort((a, b) => {
                const hA = a.height || 0;
                const hB = b.height || 0;
                if (hB !== hA) return hB - hA;
                
                const bA = a.bitrate || a.abr || 0;
                const bB = b.bitrate || b.abr || 0;
                return bB - bA;
            });
            
            const addedVideoRes = new Set();
            const addedAudioBitrate = new Set();
            
            const downloadEndpoint = window.API_ENDPOINTS?.download || '/video-downloader/api/download/';

            sortedFormats.forEach(fmt => {
                const sizeStr = (fmt.filesize && fmt.filesize > 0) ? formatBytes(fmt.filesize) : '';
                const sizeMarkup = sizeStr ? '<span class="text-surface-400 text-sm hidden sm:block ml-auto pr-4">' + sizeStr + '</span>' : '<span class="hidden sm:block ml-auto pr-4"></span>';
                const dlUrl = downloadEndpoint + '?url=' + encodeURIComponent(url) + '&format_id=' + encodeURIComponent(fmt.format_id) + '&format_type=' + encodeURIComponent(fmt.type);
                
                if (fmt.type === 'Audio Only') {
                    const rawBitrate = Math.round(fmt.bitrate || fmt.abr || 0);
                    let displayBitrate = rawBitrate;
                    if (rawBitrate >= 300) displayBitrate = 320;
                    else if (rawBitrate >= 240) displayBitrate = 256;
                    else if (rawBitrate >= 180) displayBitrate = 192;
                    else if (rawBitrate >= 110) displayBitrate = 128;
                    else if (rawBitrate >= 64) displayBitrate = rawBitrate;
                    
                    const isUnknown = displayBitrate === 0;
                    const finalDisplayStr = isUnknown ? 'Best' : displayBitrate + ' kbps';
                    
                    if ((displayBitrate >= 64 || isUnknown) && !addedAudioBitrate.has(displayBitrate)) {
                        hasAudio = true;
                        addedAudioBitrate.add(displayBitrate);
                        
                        const row = document.createElement('div');
                        row.className = 'flex items-center justify-between px-6 py-4 hover:bg-surface-50 transition-colors';
                        
                        row.innerHTML = `
                            <div class="flex items-center gap-4 sm:gap-8 flex-grow">
                                <span class="font-bold text-surface-900 w-20">${finalDisplayStr}</span>
                                <span class="font-medium text-surface-500 uppercase">MP3</span>
                                ${sizeMarkup}
                            </div>
                            <a href="${dlUrl}" class="flex-shrink-0 inline-flex items-center justify-center w-10 h-10 bg-brand-50 hover:bg-brand-600 text-brand-600 hover:text-white rounded-full transition-colors shadow-sm" title="Download MP3">
                                <i data-lucide="download" class="w-5 h-5"></i>
                            </a>
                        `;
                        audioFormatsBody.appendChild(row);
                    }
                } else if (fmt.type === 'Video + Audio') {
                    let res = fmt.height ? fmt.height + 'p' : null;
                    if (!res && fmt.resolution && fmt.resolution !== 'unknown') {
                        res = fmt.resolution;
                    }
                    
                    if (res && (/^\d+p$/.test(res) || res === 'Highest Quality') && !addedVideoRes.has(res)) {
                        hasVideo = true;
                        addedVideoRes.add(res);
                        
                        const row = document.createElement('div');
                        row.className = 'flex items-center justify-between px-6 py-4 hover:bg-surface-50 transition-colors';
                        
                        row.innerHTML = `
                            <div class="flex items-center gap-4 sm:gap-8 flex-grow">
                                <span class="font-bold text-surface-900 w-16">${res}</span>
                                <span class="font-medium text-surface-500 uppercase">MP4</span>
                                ${sizeMarkup}
                            </div>
                            <a href="${dlUrl}" class="flex-shrink-0 inline-flex items-center justify-center w-10 h-10 bg-brand-50 hover:bg-brand-600 text-brand-600 hover:text-white rounded-full transition-colors shadow-sm" title="Download MP4">
                                <i data-lucide="download" class="w-5 h-5"></i>
                            </a>
                        `;
                        videoFormatsBody.appendChild(row);
                    }
                }
            });
            
            videoFormatsContainer.classList.toggle('hidden', !hasVideo);
            audioFormatsContainer.classList.toggle('hidden', !hasAudio);
            
            const noFormatsMessage = document.getElementById('no-formats-message');
            if (noFormatsMessage) {
                noFormatsMessage.classList.toggle('hidden', hasVideo || hasAudio);
            }
            
            if(window.lucide) window.lucide.createIcons();
            
            // Show results
            loadingState.classList.add('hidden');
            resultsSection.classList.remove('hidden');
            
            // Scroll to results
            resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            
        } catch (err) {
            loadingState.classList.add('hidden');
            errorMessage.textContent = err.message;
            errorAlert.classList.remove('hidden');
        } finally {
            analyzeBtn.disabled = false;
            analyzeBtn.innerHTML = '<i data-lucide="arrow-right" class="w-5 h-5"></i><span>Start</span>';
            if(window.lucide) window.lucide.createIcons();
        }
    });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
