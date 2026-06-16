/* ═══════════════════════════════════════════════════════════
   Image to Video Editor - Production-Ready Implementation
   Handles audio editing, timeline drag-drop, image customization, and export
   ═══════════════════════════════════════════════════════════ */

(function() {
    'use strict';

    // ═══════════════════════════════════════════════════════════
    // STATE MANAGEMENT
    // ═══════════════════════════════════════════════════════════
    
    const state = {
        images: [],               // Array of image objects with metadata
        audioFile: null,          // Audio file object
        audioBlob: null,          // Audio blob data
        audioUrl: null,           // Audio object URL
        wavesurfer: null,         // Wavesurfer instance
        selectedImageIndex: -1,   // Currently selected image
        isPlaying: false,         // Playback state
        currentTime: 0,           // Current playback position
        totalDuration: 0,         // Total video duration
        zoomLevel: 50,            // Timeline zoom (pixels per second)
        undoStack: [],            // Undo history
        redoStack: [],            // Redo history
        audioSettings: {
            volume: 100,
            fadeIn: 0,
            fadeOut: 0,
            behavior: 'trim',
            syncDuration: true,
            trimStart: 0,
            trimEnd: null
        },
        videoSettings: {
            format: 'mp4',
            resolution: '1080p',
            aspectRatio: '16:9',
            fps: 30,
            transition: 'fade',
            defaultDuration: 3
        },
        imageEdits: {},           // Per-image edits map: { imageIndex: editsObject }
        sortableInstance: null    // Sortable.js instance for timeline
    };

    // ═══════════════════════════════════════════════════════════
    // DOM ELEMENTS
    // ═══════════════════════════════════════════════════════════
    
    const elements = {
        imgDropZone: document.getElementById('img-drop-zone'),
        imgInput: document.getElementById('img-input'),
        audioDropZone: document.getElementById('audio-drop-zone'),
        audioInput: document.getElementById('audio-input'),
        imageList: document.getElementById('image-list'),
        audioInfo: document.getElementById('audio-info'),
        audioFilename: document.getElementById('audio-filename'),
        previewCanvas: document.getElementById('preview-canvas'),
        previewPlaceholder: document.getElementById('preview-placeholder'),
        previewContainer: document.getElementById('preview-container'),
        timelineScroll: document.getElementById('timeline-scroll'),
        timelineRuler: document.getElementById('timeline-ruler'),
        rulerLabels: document.getElementById('ruler-labels'),
        imageClipsContainer: document.getElementById('image-clips-container'),
        waveformContainer: document.getElementById('waveform-container'),
        audioTrack: document.getElementById('audio-track'),
        imageTrack: document.getElementById('image-track'),
        playheadRuler: document.getElementById('playhead-ruler'),
        playheadAudio: document.getElementById('playhead-audio'),
        playheadImages: document.getElementById('playhead-images'),
        btnPlay: document.getElementById('btn-play'),
        btnSkipStart: document.getElementById('btn-skip-start'),
        btnZoomIn: document.getElementById('btn-zoom-in'),
        btnZoomOut: document.getElementById('btn-zoom-out'),
        btnFit: document.getElementById('btn-fit'),
        btnUndo: document.getElementById('btn-undo'),
        btnRedo: document.getElementById('btn-redo'),
        btnGenerate: document.getElementById('btn-generate'),
        btnGenerateMain: document.getElementById('btn-generate-main'),
        trimOverlay: document.getElementById('audio-trim-overlay'),
        trimRegion: document.getElementById('trim-region'),
        trimStartHandle: document.getElementById('trim-start-handle'),
        trimEndHandle: document.getElementById('trim-end-handle'),
        trimDurationLabel: document.getElementById('trim-duration-label'),
        btnPreviewTrim: document.getElementById('btn-preview-trim'),
        removeAudioBtn: document.getElementById('remove-audio'),
        settingVolume: document.getElementById('setting-volume'),
        settingFadeIn: document.getElementById('setting-fade-in'),
        settingFadeOut: document.getElementById('setting-fade-out'),
        settingAudioBehavior: document.getElementById('setting-audio-behavior'),
        settingSyncDuration: document.getElementById('setting-sync-duration'),
        settingFormat: document.getElementById('setting-format'),
        settingResolution: document.getElementById('setting-resolution'),
        settingAspect: document.getElementById('setting-aspect'),
        settingFps: document.getElementById('setting-fps'),
        settingTransition: document.getElementById('setting-transition'),
        settingDuration: document.getElementById('setting-duration'),
        editOverlayTools: document.getElementById('edit-overlay-tools'),
        editZoom: document.getElementById('edit-zoom'),
        editZoomVal: document.getElementById('edit-zoom-val'),
        btnEditCrop: document.getElementById('btn-edit-crop'),
        btnEditRotateL: document.getElementById('btn-edit-rotate-l'),
        btnEditRotateR: document.getElementById('btn-edit-rotate-r'),
        btnEditFlipH: document.getElementById('btn-edit-flip-h'),
        btnEditDone: document.getElementById('btn-edit-done'),
        ctrlZoom: document.getElementById('ctrl-zoom'),
        ctrlBrightness: document.getElementById('ctrl-brightness'),
        ctrlContrast: document.getElementById('ctrl-contrast'),
        ctrlSaturation: document.getElementById('ctrl-saturation'),
        ctrlBlur: document.getElementById('ctrl-blur'),
        ctrlOpacity: document.getElementById('ctrl-opacity'),
        btnApplyEdits: document.getElementById('btn-apply-edits'),
        panelTabs: document.querySelectorAll('.panel-tab[data-tab]'),
        rpanelTabs: document.querySelectorAll('.panel-tab[data-rtab]'),
        tabImages: document.getElementById('tab-images'),
        tabAudio: document.getElementById('tab-audio'),
        rtabSettings: document.getElementById('rtab-settings'),
        rtabEdit: document.getElementById('rtab-edit'),
        imgCount: document.getElementById('img-count'),
        totalDuration: document.getElementById('total-duration'),
        timeDisplay: document.getElementById('time-display'),
        zoomLevel: document.getElementById('zoom-level'),
        volumeValue: document.getElementById('volume-value'),
        processingModal: document.getElementById('processing-modal'),
        successModal: document.getElementById('success-modal'),
        renderProgress: document.getElementById('render-progress'),
        renderStatusText: document.getElementById('render-status-text'),
        downloadLink: document.getElementById('download-link'),
        btnNewProject: document.getElementById('btn-new-project')
    };

    // ═══════════════════════════════════════════════════════════
    // UTILITY FUNCTIONS
    // ═══════════════════════════════════════════════════════════
    
    function formatTime(seconds) {
        const mins = Math.floor(seconds / 60);
        const secs = Math.floor(seconds % 60);
        const ms = Math.floor((seconds % 1) * 10);
        return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}.${ms}`;
    }

    function showToast(message, type = 'info') {
        const toast = document.getElementById('toast');
        toast.textContent = message;
        toast.className = 'fixed bottom-6 left-1/2 -translate-x-1/2 px-4 py-2.5 rounded-xl text-xs font-semibold z-[9999] shadow-xl backdrop-blur-md pointer-events-none';
        const colors = {
            info: 'bg-indigo-600/90 text-white',
            success: 'bg-emerald-600/90 text-white',
            error: 'bg-red-600/90 text-white',
            warning: 'bg-amber-600/90 text-white'
        };
        toast.classList.add(...(colors[type] || colors.info).split(' '));
        toast.classList.remove('hidden');
        setTimeout(() => toast.classList.add('hidden'), 3000);
    }

    function saveState() {
        state.undoStack.push(JSON.stringify({
            images: state.images.map(img => ({...img})),
            audioSettings: {...state.audioSettings},
            videoSettings: {...state.videoSettings},
            imageEdits: {...state.imageEdits}
        }));
        if (state.undoStack.length > 50) state.undoStack.shift();
        state.redoStack = [];
        updateUndoRedoButtons();
    }

    function updateUndoRedoButtons() {
        elements.btnUndo.disabled = state.undoStack.length === 0;
        elements.btnRedo.disabled = state.redoStack.length === 0;
    }

    function debounce(func, wait) {
        let timeout;
        return function(...args) {
            clearTimeout(timeout);
            timeout = setTimeout(() => func(...args), wait);
        };
    }

    // ═══════════════════════════════════════════════════════════
    // IMAGE UPLOAD & MANAGEMENT
    // ═══════════════════════════════════════════════════════════
    
    function handleImageUpload(files) {
        if (state.images.length >= 100) {
            showToast('Maximum 100 images allowed', 'warning');
            return;
        }

        saveState();
        let successCount = 0;

        Array.from(files).forEach(file => {
            if (!file.type.startsWith('image/')) return;
            if (state.images.length >= 100) return;

            const reader = new FileReader();
            reader.onload = (e) => {
                const img = new Image();
                img.onload = () => {
                    const imageData = {
                        id: Date.now() + Math.random(),
                        name: file.name,
                        src: e.target.result,
                        width: img.width,
                        height: img.height,
                        duration: state.videoSettings.defaultDuration,
                        edits: {
                            zoom: 100,
                            brightness: 0,
                            contrast: 0,
                            saturation: 0,
                            blur: 0,
                            opacity: 100,
                            rotation: 0,
                            flipH: false,
                            flipV: false,
                            filter: 'none',
                            crop: null
                        }
                    };
                    state.images.push(imageData);
                    const newIndex = state.images.length - 1;
                    state.imageEdits[newIndex] = {...imageData.edits};
                    successCount++;
                    updateImageList();
                    updateTimeline();
                    recalculateTotalDuration();
                    updatePreview();
                };
                img.onerror = () => showToast(`Failed to load image: ${file.name}`, 'error');
                img.src = e.target.result;
            };
            reader.onerror = () => showToast(`Failed to read file: ${file.name}`, 'error');
            reader.readAsDataURL(file);
        });

        if (successCount > 0) {
            showToast(`Added ${successCount} image(s)`, 'success');
        }
    }

    function updateImageList() {
        elements.imageList.innerHTML = '';
        
        state.images.forEach((img, index) => {
            const div = document.createElement('div');
            div.className = `img-list-item p-2 cursor-pointer transition-all ${index === state.selectedImageIndex ? 'selected-item' : ''}`;
            div.dataset.index = index;
            
            div.innerHTML = `
                <div class="flex items-center gap-2">
                    <img src="${img.src}" class="w-12 h-12 object-cover rounded" alt="${img.name}" loading="lazy">
                    <div class="flex-1 min-w-0">
                        <p class="text-[10px] text-gray-300 truncate font-medium">${img.name}</p>
                        <p class="text-[9px] text-gray-500">${img.duration.toFixed(1)}s</p>
                    </div>
                    <button class="text-gray-500 hover:text-red-400 p-1 remove-img transition-colors" data-index="${index}" title="Remove">
                        <svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                    </button>
                </div>
            `;
            
            // Click to select image
            div.addEventListener('click', (e) => {
                if (!e.target.closest('.remove-img')) {
                    selectImage(index);
                }
            });
            
            // Remove button
            div.querySelector('.remove-img').addEventListener('click', (e) => {
                e.stopPropagation();
                removeImage(index);
            });
            
            elements.imageList.appendChild(div);
        });
        
        elements.imgCount.textContent = `${state.images.length} image${state.images.length !== 1 ? 's' : ''}`;
        const total = state.images.reduce((sum, img) => sum + img.duration, 0);
        elements.totalDuration.textContent = `${total.toFixed(1)}s total`;
        state.totalDuration = total;
    }

    function selectImage(index) {
        if (index < 0 || index >= state.images.length) {
            index = -1;
        }

        state.selectedImageIndex = index;
        updateImageList();
        
        if (index >= 0 && index < state.images.length) {
            const img = state.images[index];
            elements.editOverlayTools.classList.remove('hidden');
            elements.editOverlayTools.classList.add('flex');
            document.getElementById('editing-label').textContent = `Editing: ${img.name}`;
            
            // Update controls with current values
            const edits = img.edits || state.imageEdits[index] || {};
            elements.editZoom.value = edits.zoom || 100;
            elements.editZoomVal.textContent = `${edits.zoom || 100}%`;
            
            // Show edit panel
            document.getElementById('no-img-selected').classList.add('hidden');
            document.getElementById('img-edit-controls').classList.remove('hidden');
            
            // Update all sliders
            elements.ctrlZoom.value = edits.zoom || 100;
            elements.ctrlBrightness.value = edits.brightness || 0;
            elements.ctrlContrast.value = edits.contrast || 0;
            elements.ctrlSaturation.value = edits.saturation || 0;
            elements.ctrlBlur.value = edits.blur || 0;
            elements.ctrlOpacity.value = edits.opacity || 100;
            
            document.getElementById('val-zoom').textContent = `${edits.zoom || 100}%`;
            document.getElementById('val-brightness').textContent = edits.brightness || 0;
            document.getElementById('val-contrast').textContent = edits.contrast || 0;
            document.getElementById('val-saturation').textContent = edits.saturation || 0;
            document.getElementById('val-blur').textContent = `${edits.blur || 0}px`;
            document.getElementById('val-opacity').textContent = `${edits.opacity || 100}%`;
            
            // Update filter buttons
            document.querySelectorAll('.filter-btn').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.filter === (edits.filter || 'none'));
            });
        } else {
            elements.editOverlayTools.classList.add('hidden');
            elements.editOverlayTools.classList.remove('flex');
            document.getElementById('no-img-selected').classList.remove('hidden');
            document.getElementById('img-edit-controls').classList.add('hidden');
        }
        
        updateTimeline();
        updatePreview();
    }

    function removeImage(index) {
        saveState();
        state.images.splice(index, 1);
        
        // Reindex edits after removal
        const newEdits = {};
        Object.keys(state.imageEdits).forEach(key => {
            const k = parseInt(key);
            if (k > index) {
                newEdits[k - 1] = state.imageEdits[key];
            } else if (k < index) {
                newEdits[k] = state.imageEdits[key];
            }
        });
        state.imageEdits = newEdits;
        
        if (state.selectedImageIndex === index) {
            state.selectedImageIndex = -1;
        } else if (state.selectedImageIndex > index) {
            state.selectedImageIndex--;
        }
        
        updateImageList();
        updateTimeline();
        recalculateTotalDuration();
        updatePreview();
        showToast('Image removed', 'success');
    }

    // ═══════════════════════════════════════════════════════════
    // AUDIO UPLOAD & WAVEFORM
    // ═══════════════════════════════════════════════════════════
    
    function handleAudioUpload(file) {
        if (!file || !file.type.startsWith('audio/')) {
            showToast('Please upload an audio file', 'error');
            return;
        }
        
        saveState();
        state.audioFile = file;
        elements.audioFilename.textContent = file.name;
        elements.audioInfo.classList.remove('hidden');
        
        const reader = new FileReader();
        reader.onload = (e) => {
            state.audioBlob = e.target.result;
            state.audioUrl = URL.createObjectURL(file);
            initWaveSurfer(state.audioUrl);
        };
        reader.onerror = () => showToast('Failed to read audio file', 'error');
        reader.readAsDataURL(file);
    }

    function initWaveSurfer(url) {
        if (state.wavesurfer) {
            state.wavesurfer.destroy();
        }
        
        state.wavesurfer = WaveSurfer.create({
            container: elements.waveformContainer,
            waveColor: '#4f46e5',
            progressColor: '#818cf8',
            cursorColor: '#f43f5e',
            barWidth: 2,
            barGap: 2,
            barRadius: 2,
            height: 60,
            normalize: true,
            backend: 'WebAudio',
            interact: true
        });
        
        state.wavesurfer.load(url);
        
        state.wavesurfer.on('ready', () => {
            const duration = state.wavesurfer.getDuration();
            state.audioSettings.trimEnd = duration;
            updateTimeline();
            updateTrimRegion();
            createTrimRegion();
            
            if (state.audioSettings.syncDuration && state.images.length > 0) {
                syncDurationWithAudio();
            }
            showToast(`Audio loaded (${formatTime(duration)})`, 'success');
        });
        
        state.wavesurfer.on('audioprocess', () => {
            if (state.isPlaying) {
                const currentTime = state.wavesurfer.getCurrentTime();
                const duration = state.wavesurfer.getDuration();
                const trimEnd = state.audioSettings.trimEnd || duration;
                
                if (currentTime >= trimEnd) {
                    state.wavesurfer.pause();
                    state.isPlaying = false;
                    updatePlayPauseUI();
                    return;
                }
                
                state.currentTime = currentTime;
                updatePlayhead();
                updateTimeDisplay();
            }
        });
        
        state.wavesurfer.on('error', (err) => {
            showToast(`Audio error: ${err}`, 'error');
            logger.error('WaveSurfer error:', err);
        });
    }

    function updateTrimRegion() {
        if (!state.wavesurfer) return;
        
        const duration = state.wavesurfer.getDuration();
        const start = state.audioSettings.trimStart;
        const end = state.audioSettings.trimEnd || duration;
        
        const startPercent = (start / duration) * 100;
        const endPercent = (end / duration) * 100;
        
        elements.trimRegion.style.left = `${startPercent}%`;
        elements.trimRegion.style.width = `${endPercent - startPercent}%`;
        elements.trimStartHandle.style.left = `${startPercent}%`;
        elements.trimEndHandle.style.left = `${endPercent}%`;
        
        const trimDuration = end - start;
        elements.trimDurationLabel.textContent = formatTime(trimDuration);
    }

    function createTrimRegion() {
        if (!state.wavesurfer || !state.wavesurfer.plugins || !state.wavesurfer.plugins.regions) return;
        
        const duration = state.wavesurfer.getDuration();
        const start = state.audioSettings.trimStart;
        const end = state.audioSettings.trimEnd || duration;
        
        state.wavesurfer.plugins.regions.clear();
        
        const trimRegion = state.wavesurfer.plugins.regions.add({
            start: start,
            end: end,
            color: 'rgba(79, 70, 229, 0.3)',
            drag: true,
            resize: true
        });
        
        trimRegion.on('update-end', () => {
            state.audioSettings.trimStart = trimRegion.start;
            state.audioSettings.trimEnd = trimRegion.end;
            updateTrimRegion();
            if (state.audioSettings.syncDuration && state.images.length > 0) {
                syncDurationWithAudio();
            }
        });
    }

    function syncDurationWithAudio() {
        if (!state.wavesurfer || state.images.length === 0) return;
        
        const audioDuration = state.wavesurfer.getDuration();
        const trimDuration = (state.audioSettings.trimEnd || audioDuration) - state.audioSettings.trimStart;
        const totalImageDuration = state.images.reduce((sum, img) => sum + img.duration, 0);
        
        if (trimDuration < totalImageDuration) {
            // Compress image durations proportionally
            const ratio = trimDuration / totalImageDuration;
            state.images.forEach(img => {
                img.duration = Math.max(0.5, img.duration * ratio);
            });
        }
        
        updateImageList();
        updateTimeline();
        recalculateTotalDuration();
    }

    function recalculateTotalDuration() {
        const totalImageDuration = state.images.reduce((sum, img) => sum + img.duration, 0);
        
        if (state.wavesurfer) {
            const audioDuration = state.wavesurfer.getDuration();
            const trimDuration = (state.audioSettings.trimEnd || audioDuration) - (state.audioSettings.trimStart || 0);
            state.totalDuration = Math.max(totalImageDuration, trimDuration);
        } else {
            state.totalDuration = totalImageDuration;
        }
        
        elements.totalDuration.textContent = `${state.totalDuration.toFixed(1)}s total`;
    }

    // ═══════════════════════════════════════════════════════════
    // TIMELINE SYSTEM
    // ═══════════════════════════════════════════════════════════
    
    function updateTimeline() {
        const totalWidth = Math.max(
            elements.timelineScroll.clientWidth,
            state.totalDuration * state.zoomLevel + 200
        );
        
        elements.timelineRuler.style.width = `${totalWidth}px`;
        elements.audioTrack.style.width = `${totalWidth}px`;
        elements.imageTrack.style.width = `${totalWidth}px`;
        
        updateRuler(totalWidth);
        updateImageClips();
        
        if (state.wavesurfer) {
            elements.trimOverlay.classList.remove('hidden');
        } else {
            elements.trimOverlay.classList.add('hidden');
        }
        
        recalculateTotalDuration();
    }

    function updateRuler(width) {
        elements.rulerLabels.innerHTML = '';
        const interval = state.zoomLevel >= 100 ? 1 : (state.zoomLevel >= 50 ? 2 : 5);
        
        for (let t = 0; t <= state.totalDuration + 10; t += interval) {
            const label = document.createElement('div');
            label.className = 'absolute bottom-1 text-[9px] text-gray-500 font-mono';
            label.style.left = `${t * state.zoomLevel + 96}px`;
            label.textContent = formatTime(t);
            elements.rulerLabels.appendChild(label);
        }
    }

    function updateImageClips() {
        elements.imageClipsContainer.innerHTML = '';
        
        let currentTime = 0;
        state.images.forEach((img, index) => {
            const clip = document.createElement('div');
            clip.className = `clip-block ${index === state.selectedImageIndex ? 'selected' : ''} transition-all`;
            clip.style.left = `${currentTime * state.zoomLevel}px`;
            clip.style.width = `${img.duration * state.zoomLevel}px`;
            clip.dataset.index = index;
            clip.draggable = true;
            
            clip.innerHTML = `
                <img src="${img.src}" alt="${img.name}" loading="lazy">
                <span class="clip-label">${img.name}</span>
                <div class="resize-handle resize-handle-left" data-handle="left"></div>
                <div class="resize-handle resize-handle-right" data-handle="right"></div>
            `;
            
            clip.addEventListener('click', (e) => {
                if (!e.target.classList.contains('resize-handle')) {
                    selectImage(index);
                    seekTo(currentTime);
                }
            });
            
            elements.imageClipsContainer.appendChild(clip);
            currentTime += img.duration;
        });
        
        // Initialize Sortable.js for drag-drop reordering
        if (typeof Sortable !== 'undefined' && state.images.length > 0) {
            if (state.sortableInstance) {
                state.sortableInstance.destroy();
            }
            
            state.sortableInstance = Sortable.create(elements.imageClipsContainer, {
                animation: 150,
                ghostClass: 'clip-block-ghost',
                onEnd: (evt) => {
                    saveState();
                    const oldIndex = evt.oldIndex;
                    const newIndex = evt.newIndex;
                    
                    // Reorder images array
                    const [moved] = state.images.splice(oldIndex, 1);
                    state.images.splice(newIndex, 0, moved);
                    
                    // Reorder imageEdits map
                    const newEdits = {};
                    const imageIndices = Array.from({length: state.images.length}, (_, i) => i);
                    
                    // Create mapping of old indices to new indices
                    imageIndices.forEach((_, idx) => {
                        if (idx === oldIndex) {
                            newEdits[newIndex] = state.imageEdits[oldIndex];
                        } else if (idx < oldIndex && newIndex < idx) {
                            newEdits[idx + 1] = state.imageEdits[idx];
                        } else if (idx > oldIndex && newIndex > idx) {
                            newEdits[idx - 1] = state.imageEdits[idx];
                        } else if (idx < oldIndex || idx > newIndex) {
                            newEdits[idx] = state.imageEdits[idx];
                        }
                    });
                    
                    state.imageEdits = newEdits;
                    
                    if (state.selectedImageIndex === oldIndex) {
                        state.selectedImageIndex = newIndex;
                    } else if (state.selectedImageIndex > oldIndex && state.selectedImageIndex <= newIndex) {
                        state.selectedImageIndex--;
                    } else if (state.selectedImageIndex < oldIndex && state.selectedImageIndex >= newIndex) {
                        state.selectedImageIndex++;
                    }
                    
                    updateImageList();
                    updateTimeline();
                    updatePreview();
                    showToast('Timeline reordered', 'success');
                }
            });
        }
    }

    function updatePlayhead() {
        const position = state.currentTime * state.zoomLevel + 96;
        elements.playheadRuler.style.left = `${position}px`;
        elements.playheadAudio.style.left = `${position}px`;
        elements.playheadImages.style.left = `${position}px`;
    }

    function updateTimeDisplay() {
        elements.timeDisplay.textContent = `${formatTime(state.currentTime)} / ${formatTime(state.totalDuration)}`;
    }

    function seekTo(time) {
        if (state.wavesurfer) {
            const trimStart = state.audioSettings.trimStart || 0;
            const trimEnd = state.audioSettings.trimEnd || state.wavesurfer.getDuration();
            time = Math.max(trimStart, Math.min(time, trimEnd));
        }
        
        state.currentTime = Math.max(0, time);
        updatePlayhead();
        updateTimeDisplay();
        
        if (state.wavesurfer) {
            state.wavesurfer.seekTo(state.currentTime / state.wavesurfer.getDuration());
        }
        
        updatePreview();
    }

    // ═══════════════════════════════════════════════════════════
    // PREVIEW SYSTEM
    // ═══════════════════════════════════════════════════════════
    
    function updatePreview() {
        if (state.images.length === 0) {
            elements.previewPlaceholder.classList.remove('hidden');
            elements.previewCanvas.style.display = 'none';
            return;
        }
        
        elements.previewPlaceholder.classList.add('hidden');
        elements.previewCanvas.style.display = 'block';
        
        const ctx = elements.previewCanvas.getContext('2d');
        if (!ctx) return;
        
        const resolutions = {
            '360p': [640, 360],
            '480p': [854, 480],
            '720p': [1280, 720],
            '1080p': [1920, 1080],
            '2k': [2560, 1440],
            '4k': [3840, 2160]
        };
        
        const [width, height] = resolutions[state.videoSettings.resolution] || [1920, 1080];
        elements.previewCanvas.width = width;
        elements.previewCanvas.height = height;
        
        let totalDuration = state.images.reduce((sum, img) => sum + img.duration, 0);
        
        // Find current image
        let currentTime = 0;
        let currentImage = state.images[0];
        let currentEdits = state.imageEdits[0] || currentImage.edits || {};
        
        const effectiveTime = totalDuration > 0 ? state.currentTime % totalDuration : 0;
        
        for (let i = 0; i < state.images.length; i++) {
            if (effectiveTime >= currentTime && effectiveTime < currentTime + state.images[i].duration) {
                currentImage = state.images[i];
                currentEdits = state.imageEdits[i] || currentImage.edits || {};
                break;
            }
            currentTime += state.images[i].duration;
        }
        
        if (currentImage) {
            const img = new Image();
            img.onload = () => {
                ctx.fillStyle = '#000';
                ctx.fillRect(0, 0, width, height);
                
                ctx.save();
                ctx.translate(width / 2, height / 2);
                ctx.rotate(((currentEdits.rotation || 0) % 360) * Math.PI / 180);
                if (currentEdits.flipH) ctx.scale(-1, 1);
                if (currentEdits.flipV) ctx.scale(1, -1);
                const zoom = (currentEdits.zoom || 100) / 100;
                ctx.scale(zoom, zoom);
                ctx.translate(-width / 2, -height / 2);
                
                const brightness = (currentEdits.brightness || 0);
                const contrast = (currentEdits.contrast || 0);
                const saturation = (currentEdits.saturation || 0);
                const blur = Math.max(0, currentEdits.blur || 0);
                const opacity = Math.max(0, Math.min(1, (currentEdits.opacity || 100) / 100));
                
                let filters = `brightness(${100 + Math.max(-100, Math.min(100, brightness))}%)`;
                filters += ` contrast(${100 + Math.max(-100, Math.min(100, contrast))}%)`;
                filters += ` saturate(${100 + Math.max(-100, Math.min(100, saturation))}%)`;
                if (blur > 0) filters += ` blur(${blur}px)`;
                
                const filter = currentEdits.filter || 'none';
                if (filter === 'grayscale') filters += ' grayscale(100%)';
                else if (filter === 'sepia') filters += ' sepia(100%)';
                else if (filter === 'invert') filters += ' invert(100%)';
                else if (filter === 'warm') filters += ' sepia(30%) saturate(140%)';
                else if (filter === 'cool') filters += ' saturate(80%) hue-rotate(20deg)';
                else if (filter === 'vivid') filters += ' saturate(200%) contrast(120%)';
                else if (filter === 'fade') filters += ' brightness(120%) opacity(80%)';
                
                ctx.filter = filters;
                ctx.globalAlpha = opacity;
                
                const scale = Math.min(width / img.width, height / img.height);
                const drawWidth = img.width * scale;
                const drawHeight = img.height * scale;
                const drawX = (width - drawWidth) / 2;
                const drawY = (height - drawHeight) / 2;
                
                ctx.drawImage(img, drawX, drawY, drawWidth, drawHeight);
                ctx.restore();
            };
            img.src = currentImage.src;
        }
    }

    // ═══════════════════════════════════════════════════════════
    // PLAYBACK CONTROL
    // ═══════════════════════════════════════════════════════════
    
    function startPlayback() {
        if (state.images.length === 0) {
            showToast('No images to play', 'warning');
            return;
        }
        
        state.isPlaying = true;
        document.getElementById('icon-play').classList.add('hidden');
        document.getElementById('icon-pause').classList.remove('hidden');
        document.getElementById('play-label').textContent = 'Pause';
        
        if (state.wavesurfer) {
            const trimStart = state.audioSettings.trimStart || 0;
            const trimEnd = state.audioSettings.trimEnd || state.wavesurfer.getDuration();
            const currentTime = state.wavesurfer.getCurrentTime();
            
            if (currentTime < trimStart || currentTime >= trimEnd) {
                state.wavesurfer.seekTo(trimStart / state.wavesurfer.getDuration());
                state.currentTime = trimStart;
            }
            
            state.wavesurfer.play();
        } else {
            playVideoOnly();
        }
    }

    function stopPlayback() {
        state.isPlaying = false;
        document.getElementById('icon-play').classList.remove('hidden');
        document.getElementById('icon-pause').classList.add('hidden');
        document.getElementById('play-label').textContent = 'Preview';
        
        if (state.wavesurfer) {
            state.wavesurfer.pause();
        }
    }

    function updatePlayPauseUI() {
        if (state.isPlaying) {
            document.getElementById('icon-play').classList.add('hidden');
            document.getElementById('icon-pause').classList.remove('hidden');
            document.getElementById('play-label').textContent = 'Pause';
        } else {
            document.getElementById('icon-play').classList.remove('hidden');
            document.getElementById('icon-pause').classList.add('hidden');
            document.getElementById('play-label').textContent = 'Preview';
        }
    }

    function playVideoOnly() {
        if (!state.isPlaying) return;
        
        state.currentTime += 0.016; // ~60fps
        
        if (state.currentTime >= state.totalDuration) {
            state.currentTime = 0;
            stopPlayback();
            return;
        }
        
        updatePlayhead();
        updateTimeDisplay();
        updatePreview();
        
        requestAnimationFrame(playVideoOnly);
    }

    // ═══════════════════════════════════════════════════════════
    // ZOOM CONTROL
    // ═══════════════════════════════════════════════════════════
    
    function setZoom(level) {
        state.zoomLevel = Math.max(10, Math.min(200, level));
        elements.zoomLevel.textContent = `${state.zoomLevel}px/s`;
        updateTimeline();
        updatePlayhead();
    }

    // ═══════════════════════════════════════════════════════════
    // VIDEO GENERATION
    // ═══════════════════════════════════════════════════════════
    
    async function generateVideo() {
        if (state.images.length === 0) {
            showToast('Please add at least one image', 'error');
            return;
        }
        
        // Collect current edits for all images
        state.images.forEach((img, idx) => {
            if (!state.imageEdits[idx]) {
                state.imageEdits[idx] = {...img.edits};
            }
        });
        
        elements.processingModal.classList.remove('hidden');
        elements.renderProgress.style.width = '0%';
        elements.renderStatusText.textContent = 'Preparing files...';
        
        try {
            const formData = new FormData();
            
            // Add images in current order
            state.images.forEach((img, index) => {
                const blob = dataURLtoBlob(img.src);
                formData.append('images', blob, `image_${index}_${img.name}`);
            });
            
            // Add audio if present
            if (state.audioFile) {
                formData.append('audio', state.audioFile);
            }
            
            // Add settings
            formData.append('image_durations', JSON.stringify(state.images.map(img => img.duration)));
            formData.append('fps', state.videoSettings.fps);
            formData.append('resolution', state.videoSettings.resolution);
            formData.append('aspect_ratio', state.videoSettings.aspectRatio);
            formData.append('transition', state.videoSettings.transition);
            formData.append('audio_volume', state.audioSettings.volume / 100);
            formData.append('audio_fade_in', state.audioSettings.fadeIn);
            formData.append('audio_fade_out', state.audioSettings.fadeOut);
            formData.append('audio_behavior', state.audioSettings.behavior);
            formData.append('audio_trim_start', state.audioSettings.trimStart || 0);
            formData.append('audio_trim_end', state.audioSettings.trimEnd || '');
            formData.append('filename', `slideshow_${Date.now()}`);
            
            // Add image edits
            formData.append('image_edits', JSON.stringify(state.imageEdits));
            
            elements.renderStatusText.textContent = 'Uploading to server...';
            elements.renderProgress.style.width = '10%';
            
            const response = await fetch('/video/api/image-to-video/', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || `Server error (${response.status})`);
            }
            
            // Simulate progress
            let progress = 10;
            const progressInterval = setInterval(() => {
                progress = Math.min(90, progress + 5);
                elements.renderProgress.style.width = `${progress}%`;
            }, 200);
            
            elements.renderStatusText.textContent = 'Rendering video...';
            const blob = await response.blob();
            clearInterval(progressInterval);
            elements.renderProgress.style.width = '100%';
            
            const url = URL.createObjectURL(blob);
            elements.downloadLink.href = url;
            elements.downloadLink.download = `slideshow_${Date.now()}.mp4`;
            
            setTimeout(() => {
                elements.processingModal.classList.add('hidden');
                elements.successModal.classList.remove('hidden');
                showToast('Video generated successfully!', 'success');
            }, 500);
            
        } catch (error) {
            elements.processingModal.classList.add('hidden');
            showToast(`Generation failed: ${error.message}`, 'error');
            console.error('Generation error:', error);
        }
    }

    function dataURLtoBlob(dataURL) {
        try {
            const arr = dataURL.split(',');
            const mime = arr[0].match(/:(.*?);/)[1];
            const bstr = atob(arr[1]);
            const u8arr = new Uint8Array(bstr.length);
            for (let i = 0; i < bstr.length; i++) {
                u8arr[i] = bstr.charCodeAt(i);
            }
            return new Blob([u8arr], { type: mime });
        } catch (e) {
            console.error('Blob conversion error:', e);
            throw new Error('Failed to convert image');
        }
    }

    // ═══════════════════════════════════════════════════════════
    // EVENT LISTENERS
    // ═══════════════════════════════════════════════════════════
    
    function initEventListeners() {
        // Image upload
        elements.imgDropZone.addEventListener('click', () => elements.imgInput.click());
        elements.imgInput.addEventListener('change', (e) => {
            handleImageUpload(e.target.files);
            e.target.value = '';
        });
        
        elements.imgDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            elements.imgDropZone.classList.add('dragover');
        });
        
        elements.imgDropZone.addEventListener('dragleave', () => {
            elements.imgDropZone.classList.remove('dragover');
        });
        
        elements.imgDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            elements.imgDropZone.classList.remove('dragover');
            handleImageUpload(e.dataTransfer.files);
        });
        
        // Audio upload
        elements.audioDropZone.addEventListener('click', () => elements.audioInput.click());
        elements.audioInput.addEventListener('change', (e) => {
            handleAudioUpload(e.target.files[0]);
            e.target.value = '';
        });
        
        elements.audioDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            elements.audioDropZone.classList.add('dragover');
        });
        
        elements.audioDropZone.addEventListener('dragleave', () => {
            elements.audioDropZone.classList.remove('dragover');
        });
        
        elements.audioDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            elements.audioDropZone.classList.remove('dragover');
            handleAudioUpload(e.dataTransfer.files[0]);
        });
        
        // Remove audio
        elements.removeAudioBtn.addEventListener('click', () => {
            saveState();
            state.audioFile = null;
            state.audioBlob = null;
            if (state.audioUrl) URL.revokeObjectURL(state.audioUrl);
            if (state.wavesurfer) {
                state.wavesurfer.destroy();
                state.wavesurfer = null;
            }
            elements.audioInfo.classList.add('hidden');
            elements.trimOverlay.classList.add('hidden');
            recalculateTotalDuration();
            updateTimeline();
            showToast('Audio removed', 'success');
        });
        
        // Playback controls
        elements.btnPlay.addEventListener('click', () => {
            if (state.isPlaying) stopPlayback();
            else startPlayback();
        });
        
        elements.btnSkipStart.addEventListener('click', () => seekTo(0));
        
        // Zoom controls
        elements.btnZoomIn.addEventListener('click', () => setZoom(state.zoomLevel + 10));
        elements.btnZoomOut.addEventListener('click', () => setZoom(state.zoomLevel - 10));
        elements.btnFit.addEventListener('click', () => {
            const containerWidth = elements.timelineScroll.clientWidth - 96;
            if (state.totalDuration > 0) {
                setZoom(Math.floor(containerWidth / state.totalDuration));
            }
        });
        
        // Timeline seek
        elements.timelineRuler.addEventListener('click', (e) => {
            const rect = elements.timelineRuler.getBoundingClientRect();
            const x = e.clientX - rect.left - 96;
            const time = Math.max(0, x / state.zoomLevel);
            seekTo(time);
        });
        
        elements.imageTrack.addEventListener('click', (e) => {
            if (e.target === elements.imageTrack || e.target === elements.imageClipsContainer) {
                const rect = elements.imageTrack.getBoundingClientRect();
                const x = e.clientX - rect.left - 96;
                const time = Math.max(0, x / state.zoomLevel);
                seekTo(time);
            }
        });
        
        // Audio settings
        elements.settingVolume.addEventListener('input', (e) => {
            state.audioSettings.volume = parseInt(e.target.value);
            elements.volumeValue.textContent = `${state.audioSettings.volume}%`;
            if (state.wavesurfer) {
                state.wavesurfer.setVolume(state.audioSettings.volume / 100);
            }
        });
        
        // Video settings
        elements.settingResolution.addEventListener('change', (e) => {
            state.videoSettings.resolution = e.target.value;
            updatePreview();
        });
        
        // Image edit controls
        elements.ctrlBrightness.addEventListener('input', debounce((e) => {
            if (state.selectedImageIndex >= 0) {
                state.images[state.selectedImageIndex].edits.brightness = parseInt(e.target.value);
                document.getElementById('val-brightness').textContent = e.target.value;
                updatePreview();
            }
        }, 50));
        
        elements.ctrlContrast.addEventListener('input', debounce((e) => {
            if (state.selectedImageIndex >= 0) {
                state.images[state.selectedImageIndex].edits.contrast = parseInt(e.target.value);
                document.getElementById('val-contrast').textContent = e.target.value;
                updatePreview();
            }
        }, 50));
        
        elements.ctrlSaturation.addEventListener('input', debounce((e) => {
            if (state.selectedImageIndex >= 0) {
                state.images[state.selectedImageIndex].edits.saturation = parseInt(e.target.value);
                document.getElementById('val-saturation').textContent = e.target.value;
                updatePreview();
            }
        }, 50));
        
        // Filter buttons
        document.querySelectorAll('.filter-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                if (state.selectedImageIndex >= 0) {
                    saveState();
                    state.images[state.selectedImageIndex].edits.filter = btn.dataset.filter;
                    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
                    btn.classList.add('active');
                    updatePreview();
                }
            });
        });
        
        // Generate video
        elements.btnGenerate.addEventListener('click', generateVideo);
        elements.btnGenerateMain.addEventListener('click', generateVideo);
        
        // New project
        elements.btnNewProject.addEventListener('click', () => {
            elements.successModal.classList.add('hidden');
            location.reload();
        });
        
        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            if (e.key === ' ' && !e.target.matches('input, textarea')) {
                e.preventDefault();
                elements.btnPlay.click();
            }
            if (e.ctrlKey && e.key === 'z') {
                e.preventDefault();
                elements.btnUndo.click();
            }
            if (e.ctrlKey && e.key === 'y') {
                e.preventDefault();
                elements.btnRedo.click();
            }
        });
        
        // Window resize
        window.addEventListener('resize', debounce(() => {
            updateTimeline();
            updatePreview();
        }, 200));
    }

    // ═══════════════════════════════════════════════════════════
    // INITIALIZATION
    // ═══════════════════════════════════════════════════════════
    
    function init() {
        initEventListeners();
        updateTimeline();
        updatePreview();
        console.log('Image to Video Editor initialized');
    }

    // Start when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
