document.addEventListener('DOMContentLoaded', function() {
    const analyzeBtn = document.getElementById('analyze-btn');
    const urlInput = document.getElementById('video-url');
    const loadingState = document.getElementById('loading-state');
    const errorAlert = document.getElementById('error-alert');
    const errorMessage = document.getElementById('error-message');
    const resultsSection = document.getElementById('results-section');
    
    // Metadata elements
    const videoThumbnail = document.getElementById('video-thumbnail');
    const videoDuration = document.getElementById('video-duration');
    const videoTitle = document.getElementById('video-title');
    

    
    const mp4Btn = document.getElementById('btn-mp4-toggle');
    const mp4Menu = document.getElementById('mp4-menu');
    const mp3Btn = document.getElementById('btn-mp3-toggle');
    const mp3Menu = document.getElementById('mp3-menu');
    
    // Dropdown toggles
    if (mp4Btn && mp4Menu) {
        mp4Btn.addEventListener('click', function(e) {
            e.stopPropagation();
            mp4Menu.classList.toggle('hidden');
            if (mp3Menu) mp3Menu.classList.add('hidden');
        });
    }
    
    if (mp3Btn && mp3Menu) {
        mp3Btn.addEventListener('click', function(e) {
            e.stopPropagation();
            mp3Menu.classList.toggle('hidden');
            if (mp4Menu) mp4Menu.classList.add('hidden');
        });
    }
    
    // Close dropdowns when clicking outside
    document.addEventListener('click', function(e) {
        if (mp4Menu && !mp4Menu.classList.contains('hidden') && !mp4Menu.contains(e.target) && (!mp4Btn || !mp4Btn.contains(e.target))) {
            mp4Menu.classList.add('hidden');
        }
        if (mp3Menu && !mp3Menu.classList.contains('hidden') && !mp3Menu.contains(e.target) && (!mp3Btn || !mp3Btn.contains(e.target))) {
            mp3Menu.classList.add('hidden');
        }
    });
    
    // Escape to close
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            if (mp4Menu) mp4Menu.classList.add('hidden');
            if (mp3Menu) mp3Menu.classList.add('hidden');
        }
    });
    
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
            
            const contentType = response.headers.get('content-type');
            let data = {};
            if (contentType && contentType.indexOf('application/json') !== -1) {
                data = await response.json();
            } else {
                if (!response.ok) {
                    throw new Error(`Server returned error ${response.status} without a valid JSON message.`);
                }
            }
            
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
            
            // Clear dropdown menus
            const mp4MenuItems = document.getElementById('mp4-menu-items');
            const mp3MenuItems = document.getElementById('mp3-menu-items');
            if (mp4MenuItems) mp4MenuItems.innerHTML = '';
            if (mp3MenuItems) mp3MenuItems.innerHTML = '';
            
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
                const sizeMarkup = sizeStr ? '<span class="text-surface-400 text-xs ml-auto">' + sizeStr + '</span>' : '';
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
                    const finalDisplayStr = isUnknown ? 'High Quality' : displayBitrate + ' kbps';
                    
                    let descriptor = '';
                    if (!isUnknown) {
                        if (displayBitrate >= 256) descriptor = 'High';
                        else if (displayBitrate >= 192) descriptor = 'Standard';
                        else descriptor = 'Compact';
                    }
                    const descMarkup = descriptor ? `<span class="text-surface-400 text-xs">${descriptor}</span>` : '';
                    
                    if (!addedAudioBitrate.has(displayBitrate)) {
                        hasAudio = true;
                        addedAudioBitrate.add(displayBitrate);
                        
                        const item = document.createElement('a');
                        item.href = '#'; // Handled via JS
                        item.className = 'flex items-center gap-3 px-4 py-3 hover:bg-surface-50 transition-colors group cursor-pointer';
                        item.innerHTML = `
                            <span class="font-bold text-surface-900 group-hover:text-brand-600 transition-colors w-20">${finalDisplayStr}</span>
                            ${descMarkup}
                            ${sizeMarkup}
                            <i data-lucide="download" class="w-4 h-4 text-surface-400 group-hover:text-brand-600 transition-colors ml-2"></i>
                        `;
                        
                        item.addEventListener('click', function(e) {
                            e.preventDefault();
                            if (window.isMp3Downloading) return;
                            window.isMp3Downloading = true;
                            
                            const downloadId = Date.now().toString();
                            const finalUrl = dlUrl + '&download_id=' + downloadId;
                            
                            if (mp3Menu) mp3Menu.classList.add('hidden');
                            
                            const originalContent = mp3Btn.innerHTML;
                            const originalClasses = mp3Btn.className;
                            mp3Btn.innerHTML = '<i data-lucide="loader-2" class="w-5 h-5 animate-spin"></i><span>Downloading...</span>';
                            mp3Btn.classList.add('opacity-80', 'cursor-not-allowed', 'pointer-events-none', 'relative', 'overflow-hidden');
                            if(window.lucide) window.lucide.createIcons();
                            
                            let iframe = document.getElementById('hidden-download-iframe');
                            if (!iframe) {
                                iframe = document.createElement('iframe');
                                iframe.id = 'hidden-download-iframe';
                                iframe.style.display = 'none';
                                document.body.appendChild(iframe);
                            }
                            
                            iframe.onload = function() {
                                try {
                                    const doc = iframe.contentDocument || iframe.contentWindow.document;
                                    const text = doc.body.innerText;
                                    const data = JSON.parse(text);
                                    if (data.error || data.message || data.success === false) {
                                        clearInterval(window.mp3Timer);
                                        window.isMp3Downloading = false;
                                        mp3Btn.innerHTML = originalContent;
                                        mp3Btn.className = originalClasses;
                                        if(window.lucide) window.lucide.createIcons();
                                        errorMessage.textContent = data.error || data.message || "Download failed.";
                                        errorAlert.classList.remove('hidden');
                                    }
                                } catch(err) {
                                    clearInterval(window.mp3Timer);
                                    window.isMp3Downloading = false;
                                    mp3Btn.innerHTML = originalContent;
                                    mp3Btn.className = originalClasses;
                                    if(window.lucide) window.lucide.createIcons();
                                    errorMessage.textContent = "A network error occurred during download.";
                                    errorAlert.classList.remove('hidden');
                                }
                            };
                            
                            iframe.src = finalUrl;
                            
                            window.mp3Timer = setInterval(() => {
                                // Fetch real progress
                                fetch(`/video-downloader/api/progress/?download_id=${downloadId}`)
                                    .then(r => r.json())
                                    .then(data => {
                                        if (data && data.status) {
                                            if (data.status === 'downloading') {
                                                mp3Btn.innerHTML = `
                                                    <div class="relative z-10 flex items-center gap-2">
                                                        <i data-lucide="loader-2" class="w-5 h-5 animate-spin"></i>
                                                        <span>Downloading ${data.percent}%</span>
                                                    </div>
                                                    <div class="absolute bottom-0 left-0 h-1 bg-white/40 transition-all duration-300" style="width: ${data.percent}%"></div>
                                                `;
                                                if(window.lucide) window.lucide.createIcons();
                                            } else if (data.status === 'unknown') {
                                                // ignore, initializing
                                            } else {
                                                // Merging / Processing
                                                mp3Btn.innerHTML = `
                                                    <div class="relative z-10 flex items-center gap-2">
                                                        <i data-lucide="loader-2" class="w-5 h-5 animate-spin"></i>
                                                        <span>${data.status}</span>
                                                    </div>
                                                    <div class="absolute bottom-0 left-0 h-1 bg-white/40 transition-all duration-300" style="width: 100%"></div>
                                                `;
                                                if(window.lucide) window.lucide.createIcons();
                                            }
                                        }
                                    }).catch(() => {});

                                if (document.cookie.includes('download_started=' + downloadId)) {
                                    clearInterval(window.mp3Timer);
                                    document.cookie = 'download_started=; Max-Age=0; path=/';
                                    
                                    mp3Btn.innerHTML = `
                                        <div class="relative z-10 flex items-center gap-2">
                                            <i data-lucide="check" class="w-5 h-5"></i>
                                            <span>100%</span>
                                        </div>
                                        <div class="absolute bottom-0 left-0 h-1 bg-green-400 transition-all duration-300" style="width: 100%"></div>
                                    `;
                                    if(window.lucide) window.lucide.createIcons();
                                    
                                    setTimeout(() => {
                                        window.isMp3Downloading = false;
                                        mp3Btn.innerHTML = originalContent;
                                        mp3Btn.className = originalClasses;
                                        if(window.lucide) window.lucide.createIcons();
                                    }, 1500);
                                }
                            }, 500);
                        });
                        
                        if (mp3MenuItems) mp3MenuItems.appendChild(item);
                    }
                } else if (fmt.type === 'Video + Audio') {
                    let res = fmt.height ? fmt.height + 'p' : null;
                    if (!res && fmt.resolution && fmt.resolution !== 'unknown') {
                        res = fmt.resolution;
                    }
                    
                    if (res && (/^\d+p$/.test(res) || res === 'Highest Quality') && !addedVideoRes.has(res)) {
                        hasVideo = true;
                        addedVideoRes.add(res);
                        
                        let descriptor = '';
                        if (res === '2160p') descriptor = '4K';
                        else if (res === '1440p') descriptor = '2K';
                        else if (res === '1080p') descriptor = 'Full HD';
                        else if (res === '720p') descriptor = 'HD';
                        
                        const descMarkup = descriptor ? `<span class="text-surface-400 text-xs">${descriptor}</span>` : '';
                        
                        const item = document.createElement('a');
                        item.href = '#'; // Handled via JS
                        item.className = 'flex items-center gap-3 px-4 py-3 hover:bg-surface-50 transition-colors group cursor-pointer';
                        item.innerHTML = `
                            <span class="font-bold text-surface-900 group-hover:text-brand-600 transition-colors w-16">${res}</span>
                            ${descMarkup}
                            ${sizeMarkup}
                            <i data-lucide="download" class="w-4 h-4 text-surface-400 group-hover:text-brand-600 transition-colors ml-2"></i>
                        `;
                        
                        item.addEventListener('click', function(e) {
                            e.preventDefault();
                            if (window.isMp4Downloading) return;
                            window.isMp4Downloading = true;
                            
                            const downloadId = Date.now().toString();
                            const finalUrl = dlUrl + '&download_id=' + downloadId;
                            
                            if (mp4Menu) mp4Menu.classList.add('hidden');
                            
                            const originalContent = mp4Btn.innerHTML;
                            const originalClasses = mp4Btn.className;
                            mp4Btn.innerHTML = '<i data-lucide="loader-2" class="w-5 h-5 animate-spin"></i><span>Downloading...</span>';
                            mp4Btn.classList.add('opacity-80', 'cursor-not-allowed', 'pointer-events-none', 'relative', 'overflow-hidden');
                            if(window.lucide) window.lucide.createIcons();
                            
                            let iframe = document.getElementById('hidden-download-iframe');
                            if (!iframe) {
                                iframe = document.createElement('iframe');
                                iframe.id = 'hidden-download-iframe';
                                iframe.style.display = 'none';
                                document.body.appendChild(iframe);
                            }
                            
                            iframe.onload = function() {
                                try {
                                    const doc = iframe.contentDocument || iframe.contentWindow.document;
                                    const text = doc.body.innerText;
                                    const data = JSON.parse(text);
                                    if (data.error || data.message || data.success === false) {
                                        clearInterval(window.mp4Timer);
                                        window.isMp4Downloading = false;
                                        mp4Btn.innerHTML = originalContent;
                                        mp4Btn.className = originalClasses;
                                        if(window.lucide) window.lucide.createIcons();
                                        errorMessage.textContent = data.error || data.message || "Download failed.";
                                        errorAlert.classList.remove('hidden');
                                    }
                                } catch(err) {
                                    clearInterval(window.mp4Timer);
                                    window.isMp4Downloading = false;
                                    mp4Btn.innerHTML = originalContent;
                                    mp4Btn.className = originalClasses;
                                    if(window.lucide) window.lucide.createIcons();
                                    errorMessage.textContent = "A network error occurred during download.";
                                    errorAlert.classList.remove('hidden');
                                }
                            };
                            
                            iframe.src = finalUrl;
                            
                            window.mp4Timer = setInterval(() => {
                                // Fetch real progress
                                fetch(`/video-downloader/api/progress/?download_id=${downloadId}`)
                                    .then(r => r.json())
                                    .then(data => {
                                        if (data && data.status) {
                                            if (data.status === 'downloading') {
                                                mp4Btn.innerHTML = `
                                                    <div class="relative z-10 flex items-center gap-2">
                                                        <i data-lucide="loader-2" class="w-5 h-5 animate-spin"></i>
                                                        <span>Downloading ${data.percent}%</span>
                                                    </div>
                                                    <div class="absolute bottom-0 left-0 h-1 bg-white/40 transition-all duration-300" style="width: ${data.percent}%"></div>
                                                `;
                                                if(window.lucide) window.lucide.createIcons();
                                            } else if (data.status === 'unknown') {
                                                // ignore, initializing
                                            } else {
                                                // Merging / Processing
                                                mp4Btn.innerHTML = `
                                                    <div class="relative z-10 flex items-center gap-2">
                                                        <i data-lucide="loader-2" class="w-5 h-5 animate-spin"></i>
                                                        <span>${data.status}</span>
                                                    </div>
                                                    <div class="absolute bottom-0 left-0 h-1 bg-white/40 transition-all duration-300" style="width: 100%"></div>
                                                `;
                                                if(window.lucide) window.lucide.createIcons();
                                            }
                                        }
                                    }).catch(() => {});

                                if (document.cookie.includes('download_started=' + downloadId)) {
                                    clearInterval(window.mp4Timer);
                                    document.cookie = 'download_started=; Max-Age=0; path=/';
                                    
                                    mp4Btn.innerHTML = `
                                        <div class="relative z-10 flex items-center gap-2">
                                            <i data-lucide="check" class="w-5 h-5"></i>
                                            <span>100%</span>
                                        </div>
                                        <div class="absolute bottom-0 left-0 h-1 bg-green-400 transition-all duration-300" style="width: 100%"></div>
                                    `;
                                    if(window.lucide) window.lucide.createIcons();
                                    
                                    setTimeout(() => {
                                        window.isMp4Downloading = false;
                                        mp4Btn.innerHTML = originalContent;
                                        mp4Btn.className = originalClasses;
                                        if(window.lucide) window.lucide.createIcons();
                                    }, 1500);
                                }
                            }, 500);
                        });
                        
                        if (mp4MenuItems) mp4MenuItems.appendChild(item);
                    }
                }
            });
            
            // Toggle dropdowns UI
            const mp4Container = document.getElementById('mp4-dropdown-container');
            const mp3Container = document.getElementById('mp3-dropdown-container');
            
            if (mp4Container) mp4Container.classList.toggle('hidden', !hasVideo);
            if (mp3Container) mp3Container.classList.toggle('hidden', !hasAudio);
            
            const noFormatsMessage = document.getElementById('no-formats-message');
            if (noFormatsMessage) {
                noFormatsMessage.classList.toggle('hidden', hasVideo || hasAudio);
            }
            
            try {
                if(window.lucide) window.lucide.createIcons();
            } catch(e) {
                console.warn('Lucide icons creation failed', e);
            }
            
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
});
