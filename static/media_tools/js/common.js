/**
 * MEDIA TOOLS — SHARED JAVASCRIPT ENGINE
 *
 * Handles:
 * - Video upload
 * - Click upload
 * - Drag & drop
 * - Video preview
 * - Video metadata
 * - Custom playback controls
 * - AJAX processing
 * - Processing overlay
 * - Result handling
 * - Toast notifications
 */

(function () {
    "use strict";


    /* ========================================================================
       MEDIA TOOLS GLOBAL
       ======================================================================== */

    window.MediaTools = {


        /* ====================================================================
           FORMAT TIME
           ==================================================================== */

        formatTime: function (seconds) {

            if (
                isNaN(seconds) ||
                seconds < 0 ||
                !isFinite(seconds)
            ) {
                seconds = 0;
            }

            var mins = Math.floor(seconds / 60);
            var secs = Math.floor(seconds % 60);
            var ms = Math.floor((seconds % 1) * 10);

            return (
                String(mins).padStart(2, "0") +
                ":" +
                String(secs).padStart(2, "0") +
                "." +
                ms
            );
        },


        /* ====================================================================
           FORMAT BYTES
           ==================================================================== */

        formatBytes: function (bytes) {

            if (
                !bytes ||
                bytes <= 0
            ) {
                return "0 B";
            }

            var k = 1024;

            var sizes = [
                "B",
                "KB",
                "MB",
                "GB"
            ];

            var i = Math.floor(
                Math.log(bytes) / Math.log(k)
            );

            i = Math.min(
                i,
                sizes.length - 1
            );

            return (
                bytes / Math.pow(k, i)
            ).toFixed(1) +
            " " +
            sizes[i];
        },


        /* ====================================================================
           TOAST
           ==================================================================== */

        showToast: function (
            message,
            type
        ) {

            type = type || "info";

            var existing =
                document.getElementById(
                    "mtToastNotification"
                );

            if (!existing) {

                existing =
                    document.createElement(
                        "div"
                    );

                existing.id =
                    "mtToastNotification";

                existing.className =
                    "mt-toast";

                document.body.appendChild(
                    existing
                );
            }

            existing.className =
                "mt-toast" +
                (
                    type === "error"
                        ? " error"
                        : ""
                );

            existing.textContent =
                message || "Something went wrong.";

            existing.classList.add(
                "show"
            );

            if (window._mtToastTimer) {

                clearTimeout(
                    window._mtToastTimer
                );
            }

            window._mtToastTimer =
                setTimeout(function () {

                    existing.classList.remove(
                        "show"
                    );

                }, 5000);
        },


        /* ====================================================================
           UPLOAD CARD
           ==================================================================== */

        initUploadCard: function (
            options
        ) {

            options = options || {};


            /*
             * ---------------------------------------------------------------
             * Find elements
             * ---------------------------------------------------------------
             */

            var fileInput =
                document.querySelector(
                    options.fileInputSelector ||
                    "#id_video"
                );

            var dropzone =
                document.querySelector(
                    options.dropzoneSelector ||
                    "#mtUploadDropzone"
                );

            var uploadCard =
                document.querySelector(
                    options.uploadCardSelector ||
                    "#mtUploadCard"
                );

            var editorShell =
                document.querySelector(
                    options.editorShellSelector ||
                    "#mtEditorShell"
                );

            var videoPlayer =
                document.querySelector(
                    options.videoPlayerSelector ||
                    "#mtVideoPreview"
                );

            var fileChip =
                document.querySelector(
                    "#mtSelectedFileChip"
                );

            var fileNameEl =
                document.querySelector(
                    "#mtFileName"
                );

            var fileMetaEl =
                document.querySelector(
                    "#mtFileMeta"
                );

            var removeBtn =
                document.querySelector(
                    "#mtRemoveBtn"
                );

            var changeBtn =
                document.querySelector(
                    "#mtTopChangeBtn"
                );

            var uploadTrigger =
                document.querySelector(
                    "#mtUploadTriggerBtn"
                );


            /*
             * ---------------------------------------------------------------
             * Safety check
             * ---------------------------------------------------------------
             */

            if (!fileInput) {

                console.warn(
                    "MediaTools: #id_video was not found."
                );

                return;
            }


            /*
             * ---------------------------------------------------------------
             * Prevent duplicate initialization
             * ---------------------------------------------------------------
             */

            if (
                fileInput.dataset.mtInitialized ===
                "true"
            ) {

                return;
            }

            fileInput.dataset.mtInitialized =
                "true";


            /*
             * ---------------------------------------------------------------
             * Active object URL
             * ---------------------------------------------------------------
             */

            function revokeObjectUrl() {

                if (
                    window._mtActiveObjectUrl
                ) {

                    try {

                        URL.revokeObjectURL(
                            window._mtActiveObjectUrl
                        );

                    } catch (error) {

                        console.warn(
                            "Unable to revoke video object URL.",
                            error
                        );
                    }

                    window._mtActiveObjectUrl =
                        null;
                }
            }


            /*
             * ---------------------------------------------------------------
             * Validate video file
             * ---------------------------------------------------------------
             */

            function isValidVideoFile(
                file
            ) {

                if (!file) {
                    return false;
                }

                /*
                 * Browser MIME type.
                 */

                if (
                    file.type &&
                    (file.type.indexOf("video/") === 0 || file.type.indexOf("image/gif") === 0)
                ) {
                    return true;
                }

                var filename = String(file.name || "").toLowerCase();
                var allowedExtensions = [
                    ".mp4", ".mov", ".webm", ".avi", ".mkv", ".m4v",
                    ".gif", ".wmv", ".flv", ".ts", ".3gp", ".mpeg", ".mpg"
                ];

                return allowedExtensions.some(function (extension) {
                    return filename.endsWith(extension);
                });
            }


            /*
             * ---------------------------------------------------------------
             * Handle selected file
             * ---------------------------------------------------------------
             */

            function handleFile(
                file
            ) {

                if (!file) {
                    return;
                }


                /*
                 * Validate.
                 */

                if (
                    !isValidVideoFile(file)
                ) {

                    MediaTools.showToast(
                        "Please select a valid video file.",
                        "error"
                    );

                    fileInput.value = "";

                    return;
                }


                /*
                 * 500 MB validation.
                 *
                 * This is only client-side validation.
                 * Django should also validate server-side.
                 */

                var maxSize =
                    500 *
                    1024 *
                    1024;

                if (
                    file.size > maxSize
                ) {

                    MediaTools.showToast(
                        "Video file must be 500 MB or smaller.",
                        "error"
                    );

                    fileInput.value = "";

                    return;
                }


                /*
                 * -----------------------------------------------------------
                 * File information
                 * -----------------------------------------------------------
                 */

                if (fileNameEl) {

                    fileNameEl.textContent =
                        file.name;
                }

                if (fileMetaEl) {

                    fileMetaEl.textContent =
                        MediaTools.formatBytes(
                            file.size
                        );
                }

                if (fileChip) {

                    fileChip.style.display =
                        "flex";
                }


                /*
                 * -----------------------------------------------------------
                 * Create preview URL
                 * -----------------------------------------------------------
                 */

                revokeObjectUrl();

                var blobUrl =
                    URL.createObjectURL(
                        file
                    );

                window._mtActiveObjectUrl =
                    blobUrl;


                /*
                 * -----------------------------------------------------------
                 * Video preview
                 * -----------------------------------------------------------
                 */

                if (videoPlayer) {

                    videoPlayer.pause();

                    videoPlayer.src =
                        blobUrl;

                    videoPlayer.load();


                    /*
                     * Metadata.
                     */

                    videoPlayer.onloadedmetadata =
                        function () {

                            var width =
                                videoPlayer.videoWidth ||
                                0;

                            var height =
                                videoPlayer.videoHeight ||
                                0;

                            var duration =
                                videoPlayer.duration ||
                                0;


                            /*
                             * Resolution
                             */

                            var resPill =
                                document.querySelector(
                                    ".mt-meta-res"
                                );

                            if (resPill) {

                                if (
                                    width &&
                                    height
                                ) {

                                    resPill.textContent =
                                        width +
                                        "x" +
                                        height;

                                } else {

                                    resPill.textContent =
                                        "--x--";
                                }
                            }


                            /*
                             * Duration
                             */

                            var durPill =
                                document.querySelector(
                                    ".mt-meta-dur"
                                );

                            if (durPill) {

                                durPill.textContent =
                                    MediaTools.formatTime(
                                        duration
                                    );
                            }


                            /*
                             * Reset playbar.
                             */

                            var progress =
                                document.querySelector(
                                    ".mt-scrubber-progress"
                                );

                            if (progress) {

                                progress.style.width =
                                    "0%";
                            }


                            var timeDisplay =
                                document.querySelector(
                                    ".mt-time-display"
                                );

                            if (timeDisplay) {

                                timeDisplay.textContent =
                                    "00:00.0 / " +
                                    MediaTools.formatTime(
                                        duration
                                    );
                            }


                            /*
                             * Optional callback.
                             */

                            if (
                                typeof options.onVideoLoaded ===
                                "function"
                            ) {

                                options.onVideoLoaded(
                                    videoPlayer,
                                    file
                                );
                            }
                        };


                    /*
                     * Video loading error.
                     */

                    videoPlayer.onerror =
                        function () {

                            MediaTools.showToast(
                                "Unable to preview this video file.",
                                "error"
                            );
                        };
                }


                /*
                 * -----------------------------------------------------------
                 * Hide upload card
                 * -----------------------------------------------------------
                 */

                if (uploadCard) {

                    uploadCard.style.display =
                        "none";
                }


                /*
                 * -----------------------------------------------------------
                 * Show editor
                 * -----------------------------------------------------------
                 */

                if (editorShell) {

                    editorShell.style.display =
                        "flex";
                }
            }


            /*
             * ---------------------------------------------------------------
             * File input change
             * ---------------------------------------------------------------
             */

            fileInput.addEventListener(
                "change",
                function () {

                    if (
                        this.files &&
                        this.files.length > 0
                    ) {

                        handleFile(
                            this.files[0]
                        );
                    }
                }
            );


            /*
             * ---------------------------------------------------------------
             * Upload button
             * ---------------------------------------------------------------
             */

            if (uploadTrigger) {

                uploadTrigger.addEventListener(
                    "click",
                    function (event) {

                        event.preventDefault();

                        event.stopPropagation();

                        fileInput.click();
                    }
                );
            }


            /*
             * ---------------------------------------------------------------
             * Dropzone click
             *
             * IMPORTANT:
             * Do not trigger twice when clicking
             * the button inside the dropzone.
             * ---------------------------------------------------------------
             */

            if (dropzone) {

                dropzone.addEventListener(
                    "click",
                    function (event) {

                        if (
                            event.target.closest(
                                "#mtUploadTriggerBtn"
                            )
                        ) {

                            return;
                        }

                        fileInput.click();
                    }
                );


                /*
                 * Drag enter / over
                 */

                [
                    "dragenter",
                    "dragover"
                ].forEach(
                    function (eventName) {

                        dropzone.addEventListener(
                            eventName,
                            function (event) {

                                event.preventDefault();

                                event.stopPropagation();

                                dropzone.classList.add(
                                    "drag-active"
                                );
                            }
                        );
                    }
                );


                /*
                 * Drag leave / drop
                 */

                [
                    "dragleave",
                    "drop"
                ].forEach(
                    function (eventName) {

                        dropzone.addEventListener(
                            eventName,
                            function (event) {

                                event.preventDefault();

                                event.stopPropagation();

                                dropzone.classList.remove(
                                    "drag-active"
                                );
                            }
                        );
                    }
                );


                /*
                 * Drop file
                 */

                dropzone.addEventListener(
                    "drop",
                    function (event) {

                        var files =
                            event.dataTransfer &&
                            event.dataTransfer.files;

                        if (
                            !files ||
                            !files.length
                        ) {

                            return;
                        }


                        var file =
                            files[0];


                        /*
                         * Set the actual input files.
                         *
                         * DataTransfer is used because
                         * assigning FileList directly is
                         * not reliable across browsers.
                         */

                        try {

                            var dataTransfer =
                                new DataTransfer();

                            dataTransfer.items.add(
                                file
                            );

                            fileInput.files =
                                dataTransfer.files;

                        } catch (error) {

                            console.warn(
                                "Unable to assign dropped file to input.",
                                error
                            );
                        }


                        handleFile(
                            file
                        );
                    }
                );
            }


            /*
             * ---------------------------------------------------------------
             * Reset upload/editor
             * ---------------------------------------------------------------
             */

            function resetToUpload() {

                /*
                 * Stop video.
                 */

                if (videoPlayer) {

                    try {
                        videoPlayer.pause();
                    } catch (error) {
                        // Ignore.
                    }

                    videoPlayer.removeAttribute(
                        "src"
                    );

                    videoPlayer.load();
                }


                /*
                 * Revoke object URL.
                 */

                revokeObjectUrl();


                /*
                 * Clear input.
                 */

                fileInput.value = "";


                /*
                 * Hide selected file.
                 */

                if (fileChip) {

                    fileChip.style.display =
                        "none";
                }


                /*
                 * Reset metadata.
                 */

                var resPill =
                    document.querySelector(
                        ".mt-meta-res"
                    );

                if (resPill) {

                    resPill.textContent =
                        "--x--";
                }


                var durPill =
                    document.querySelector(
                        ".mt-meta-dur"
                    );

                if (durPill) {

                    durPill.textContent =
                        "00:00.0";
                }


                /*
                 * Reset editor.
                 */

                if (editorShell) {

                    editorShell.style.display =
                        "none";
                }


                /*
                 * Show upload.
                 */

                if (uploadCard) {

                    uploadCard.style.display =
                        "block";
                }


                /*
                 * Hide result.
                 */

                var resultCard =
                    document.querySelector(
                        ".mt-result-card"
                    );

                if (resultCard) {

                    resultCard.classList.remove(
                        "show"
                    );
                }


                /*
                 * Clear result video.
                 */

                var resultVideo =
                    document.querySelector(
                        ".mt-result-preview-video"
                    );

                if (resultVideo) {

                    resultVideo.pause();

                    resultVideo.removeAttribute(
                        "src"
                    );

                    resultVideo.load();
                }


                /*
                 * Clear download link.
                 */

                var downloadBtn =
                    document.querySelector(
                        ".mt-download-btn"
                    );

                if (downloadBtn) {

                    downloadBtn.removeAttribute(
                        "href"
                    );

                    downloadBtn.removeAttribute(
                        "download"
                    );
                }
            }


            /*
             * Remove button.
             */

            if (removeBtn) {

                removeBtn.addEventListener(
                    "click",
                    function (event) {

                        event.preventDefault();

                        resetToUpload();
                    }
                );
            }


            /*
             * Change Video button.
             */

            if (changeBtn) {

                changeBtn.addEventListener(
                    "click",
                    function (event) {

                        event.preventDefault();

                        resetToUpload();
                    }
                );
            }
        },


        /* ====================================================================
           PLAYBACK BAR
           ==================================================================== */

        initPlaybar: function (
            videoPlayer
        ) {

            if (!videoPlayer) {
                return;
            }


            var playBtn =
                document.querySelector(
                    ".mt-play-pause-btn"
                );

            var scrubberTrack =
                document.querySelector(
                    ".mt-scrubber-track"
                );

            var scrubberProgress =
                document.querySelector(
                    ".mt-scrubber-progress"
                );

            var timeDisplay =
                document.querySelector(
                    ".mt-time-display"
                );


            /*
             * ---------------------------------------------------------------
             * Update play icon
             * ---------------------------------------------------------------
             */

            function updatePlayButton() {

                if (!playBtn) {
                    return;
                }

                if (
                    videoPlayer.paused ||
                    videoPlayer.ended
                ) {

                    playBtn.innerHTML =
                        '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="currentColor">' +
                        '<polygon points="5 3 19 12 5 21 5 3"></polygon>' +
                        '</svg>';

                } else {

                    playBtn.innerHTML =
                        '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="currentColor">' +
                        '<rect x="6" y="4" width="4" height="16"></rect>' +
                        '<rect x="14" y="4" width="4" height="16"></rect>' +
                        '</svg>';
                }
            }


            /*
             * ---------------------------------------------------------------
             * Toggle play
             * ---------------------------------------------------------------
             */

            function togglePlay() {

                if (
                    videoPlayer.paused ||
                    videoPlayer.ended
                ) {

                    var playPromise =
                        videoPlayer.play();

                    if (
                        playPromise &&
                        typeof playPromise.catch ===
                        "function"
                    ) {

                        playPromise.catch(
                            function (error) {

                                console.warn(
                                    "Video playback failed.",
                                    error
                                );
                            }
                        );
                    }

                } else {

                    videoPlayer.pause();
                }
            }


            if (playBtn) {

                playBtn.addEventListener(
                    "click",
                    togglePlay
                );
            }


            videoPlayer.addEventListener(
                "play",
                updatePlayButton
            );

            videoPlayer.addEventListener(
                "pause",
                updatePlayButton
            );

            videoPlayer.addEventListener(
                "ended",
                updatePlayButton
            );


            /*
             * ---------------------------------------------------------------
             * Time update
             * ---------------------------------------------------------------
             */

            videoPlayer.addEventListener(
                "timeupdate",
                function () {

                    var duration =
                        videoPlayer.duration;

                    if (
                        !duration ||
                        !isFinite(duration)
                    ) {

                        return;
                    }

                    var current =
                        videoPlayer.currentTime;

                    var pct =
                        (
                            current /
                            duration
                        ) *
                        100;


                    if (scrubberProgress) {

                        scrubberProgress.style.width =
                            pct + "%";
                    }


                    if (timeDisplay) {

                        timeDisplay.textContent =
                            MediaTools.formatTime(
                                current
                            ) +
                            " / " +
                            MediaTools.formatTime(
                                duration
                            );
                    }
                }
            );


            /*
             * ---------------------------------------------------------------
             * Scrubber
             * ---------------------------------------------------------------
             */

            if (scrubberTrack) {

                scrubberTrack.addEventListener(
                    "click",
                    function (event) {

                        var duration =
                            videoPlayer.duration;

                        if (
                            !duration ||
                            !isFinite(duration)
                        ) {

                            return;
                        }

                        var rect =
                            scrubberTrack.getBoundingClientRect();

                        var clickPos =
                            (
                                event.clientX -
                                rect.left
                            ) /
                            rect.width;

                        clickPos =
                            Math.max(
                                0,
                                Math.min(
                                    1,
                                    clickPos
                                )
                            );

                        videoPlayer.currentTime =
                            clickPos *
                            duration;
                    }
                );
            }


            /*
             * Initial icon.
             */

            updatePlayButton();
        },


        /* ====================================================================
           AJAX FORM
           ==================================================================== */

        initAjaxForm: function (
            formSelector,
            options
        ) {

            options =
                options || {};


            var form =
                document.querySelector(
                    formSelector
                );

            if (!form) {

                console.warn(
                    "MediaTools: form not found:",
                    formSelector
                );

                return;
            }


            /*
             * Prevent duplicate initialization.
             */

            if (
                form.dataset.mtAjaxInitialized ===
                "true"
            ) {

                return;
            }

            form.dataset.mtAjaxInitialized =
                "true";


            var processBtn =
                form.querySelector(
                    'button[type="submit"]'
                );

            var processingOverlay =
                document.querySelector(
                    ".mt-processing-overlay"
                );

            var resultCard =
                document.querySelector(
                    ".mt-result-card"
                );

            var resultVideo =
                document.querySelector(
                    ".mt-result-preview-video"
                );

            var downloadBtn =
                document.querySelector(
                    ".mt-download-btn"
                );


            /*
             * ---------------------------------------------------------------
             * Submit
             * ---------------------------------------------------------------
             */

            form.addEventListener(
                "submit",
                function (event) {

                    event.preventDefault();


                    /*
                     * -------------------------------------------------------
                     * Check video
                     * -------------------------------------------------------
                     */

                    var fileInputs =
                        form.querySelectorAll(
                            'input[type="file"]'
                        );

                    var hasSelectedFiles = false;
                    fileInputs.forEach(function (fi) {
                        if (fi.files && fi.files.length > 0) {
                            hasSelectedFiles = true;
                        }
                    });

                    if (!hasSelectedFiles) {
                        var fallbackInput = document.querySelector("#id_video, #id_videos_file, input[type=\"file\"]");
                        if (fallbackInput && fallbackInput.files && fallbackInput.files.length > 0) {
                            hasSelectedFiles = true;
                        }
                    }

                    if (!hasSelectedFiles) {

                        MediaTools.showToast(
                            "Please choose a video file first.",
                            "error"
                        );

                        return;
                    }


                    /*
                     * -------------------------------------------------------
                     * Check form validity
                     * -------------------------------------------------------
                     */

                    if (
                        !form.checkValidity()
                    ) {

                        form.reportValidity();

                        return;
                    }


                    /*
                     * -------------------------------------------------------
                     * Clear previous result
                     * -------------------------------------------------------
                     */

                    if (resultCard) {

                        resultCard.classList.remove(
                            "show"
                        );
                    }


                    if (downloadBtn) {

                        downloadBtn.removeAttribute(
                            "href"
                        );

                        downloadBtn.removeAttribute(
                            "download"
                        );
                    }


                    if (resultVideo) {

                        resultVideo.pause();

                        resultVideo.removeAttribute(
                            "src"
                        );

                        resultVideo.load();
                    }


                    /*
                     * -------------------------------------------------------
                     * Show processing
                     * -------------------------------------------------------
                     */

                    if (processingOverlay) {

                        processingOverlay.classList.add(
                            "show"
                        );
                    }


                    if (processBtn) {

                        processBtn.disabled =
                            true;
                    }


                    /*
                     * -------------------------------------------------------
                     * Form data
                     * -------------------------------------------------------
                     */

                    var formData =
                        new FormData(form);


                    /*
                     * -------------------------------------------------------
                     * POST
                     * -------------------------------------------------------
                     */

                    fetch(
                        form.action ||
                        window.location.href,
                        {
                            method: "POST",

                            headers: {
                                "X-Requested-With":
                                    "XMLHttpRequest"
                            },

                            body: formData
                        }
                    )
                    .then(
                        function (response) {

                            return response
                                .text()
                                .then(
                                    function (text) {

                                        var data;

                                        try {

                                            data =
                                                JSON.parse(
                                                    text
                                                );

                                        } catch (
                                            parseError
                                        ) {

                                            throw new Error(
                                                "Server returned an invalid response."
                                            );
                                        }

                                        return {
                                            status:
                                                response.status,

                                            ok:
                                                response.ok,

                                            data:
                                                data
                                        };
                                    }
                                );
                        }
                    )
                    .then(
                        function (responseObject) {

                            if (
                                processingOverlay
                            ) {

                                processingOverlay.classList.remove(
                                    "show"
                                );
                            }


                            if (processBtn) {

                                processBtn.disabled =
                                    false;
                            }


                            var data =
                                responseObject.data;


                            /*
                             * ------------------------------------------------
                             * Success
                             * ------------------------------------------------
                             */

                            if (
                                responseObject.ok &&
                                data &&
                                data.success
                            ) {

                                var newUrl =
                                    data.video_url ||
                                    data.download_url;


                                if (!newUrl) {

                                    throw new Error(
                                        "Server did not return a video URL."
                                    );
                                }


                                /*
                                 * Download
                                 */

                                if (
                                    downloadBtn
                                ) {

                                    downloadBtn.href =
                                        newUrl;

                                    downloadBtn.setAttribute(
                                        "download",
                                        data.filename ||
                                        "processed_video.mp4"
                                    );
                                }


                                /*
                                 * Result preview
                                 */

                                if (
                                    resultVideo
                                ) {

                                    resultVideo.src =
                                        newUrl;

                                    resultVideo.load();
                                }


                                /*
                                 * Result card
                                 */

                                if (
                                    resultCard
                                ) {

                                    resultCard.classList.add(
                                        "show"
                                    );
                                }


                                MediaTools.showToast(
                                    "Video ready for download!",
                                    "success"
                                );


                                /*
                                 * Scroll result into view.
                                 */

                                if (
                                    resultCard
                                ) {

                                    setTimeout(
                                        function () {

                                            resultCard.scrollIntoView(
                                                {
                                                    behavior:
                                                        "smooth",

                                                    block:
                                                        "nearest"
                                                }
                                            );

                                        },
                                        100
                                    );
                                }


                                /*
                                 * Optional callback.
                                 */

                                if (
                                    typeof options.onSuccess ===
                                    "function"
                                ) {

                                    options.onSuccess(
                                        data
                                    );
                                }


                                return;
                            }


                            /*
                             * ------------------------------------------------
                             * Server error
                             * ------------------------------------------------
                             */

                            var errorMessage =
                                (
                                    data &&
                                    (
                                        data.message ||
                                        data.error
                                    )
                                ) ||
                                "Processing failed. Please check your video settings.";


                            MediaTools.showToast(
                                errorMessage,
                                "error"
                            );
                        }
                    )
                    .catch(
                        function (error) {

                            console.error(
                                "MediaTools AJAX error:",
                                error
                            );


                            if (
                                processingOverlay
                            ) {

                                processingOverlay.classList.remove(
                                    "show"
                                );
                            }


                            if (processBtn) {

                                processBtn.disabled =
                                    false;
                            }


                            MediaTools.showToast(
                                error.message ||
                                "An unexpected error occurred. Please try again.",
                                "error"
                            );
                        }
                    );
                }
            );
        }
    };


    /* =========================================================================
       DOM READY
       ========================================================================= */

    document.addEventListener(
        "DOMContentLoaded",
        function () {

            /*
             * We intentionally do not automatically
             * initialize everything here.
             *
             * Each template calls:
             *
             * MediaTools.initUploadCard(...)
             * MediaTools.initPlaybar(...)
             * MediaTools.initAjaxForm(...)
             *
             * This keeps the common engine compatible
             * with all 31 templates.
             */

        }
    );

})();
