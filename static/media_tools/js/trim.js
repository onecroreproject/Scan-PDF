document.addEventListener("DOMContentLoaded", () => {

    "use strict";


    /* =========================================================
       ELEMENTS
    ========================================================= */

    const uploadScreen =
        document.getElementById("uploadScreen");

    const editorScreen =
        document.getElementById("editorScreen");

    const exportingScreen =
        document.getElementById("exportingScreen");

    const resultMode =
        document.getElementById("resultMode");


    const uploadButton =
        document.getElementById("uploadButton");


    const trimForm =
        document.getElementById("trimForm");


    const videoInput =
        trimForm?.querySelector(
            'input[name="video"]'
        );


    const startTimeInput =
        trimForm?.querySelector(
            'input[name="start_time"]'
        );


    const endTimeInput =
        trimForm?.querySelector(
            'input[name="end_time"]'
        );


    const backgroundInput =
        trimForm?.querySelector(
            'input[name="background_color"]'
        );


    const outputFormatInput =
        trimForm?.querySelector(
            'input[name="output_format"]'
        );


    const sourceVideo =
        document.getElementById("sourceVideo");


    const resultVideo =
        document.getElementById("resultVideo");


    /* TIME */

    const startTimeDisplay =
        document.getElementById("startTimeDisplay");

    const endTimeDisplay =
        document.getElementById("endTimeDisplay");

    const durationDisplay =
        document.getElementById("durationDisplay");

    const videoDurationDisplay =
        document.getElementById("videoDurationDisplay");


    /* TIMELINE */

    const timeline =
        document.getElementById("timeline");

    const timelineSelection =
        document.getElementById("timelineSelection");

    const timelinePlayhead =
        document.getElementById("timelinePlayhead");

    const startHandle =
        document.getElementById("startHandle");

    const endHandle =
        document.getElementById("endHandle");

    const timelineStartLabel =
        document.getElementById("timelineStartLabel");

    const timelineEndLabel =
        document.getElementById("timelineEndLabel");

    const timelineSelectedValue =
        document.getElementById("timelineSelectedValue");


    /* EXPORT */

    const exportButton =
        document.getElementById("exportButton");


    /* RESULT */

    const downloadButton =
        document.getElementById("downloadButton");

    const editAgainButton =
        document.getElementById("editAgainButton");

    const resultStartDisplay =
        document.getElementById("resultStartDisplay");

    const resultEndDisplay =
        document.getElementById("resultEndDisplay");

    const resultDurationDisplay =
        document.getElementById("resultDurationDisplay");

    const resultFormatDisplay =
        document.getElementById("resultFormatDisplay");


    /* PLAYLIST */

    const playlistInput =
        document.getElementById("playlistInput");

    const addPlaylistButton =
        document.getElementById("addPlaylistButton");

    const playlist =
        document.getElementById("playlist");

    const playlistCount =
        document.getElementById("playlistCount");


    const currentVideoName =
        document.getElementById("currentVideoName");

    const videoCurrentName =
        document.getElementById("videoCurrentName");

    const videoResolution =
        document.getElementById("videoResolution");


    /* PLAYBACK */

    const previousVideoButton =
        document.getElementById("previousVideoButton");

    const nextVideoButton =
        document.getElementById("nextVideoButton");

    const playPauseButton =
        document.getElementById("playPauseButton");


    const previousVideoButtonTop =
        document.getElementById("previousVideoButtonTop");

    const nextVideoButtonTop =
        document.getElementById("nextVideoButtonTop");

    const playPauseButtonTop =
        document.getElementById("playPauseButtonTop");


    const playerStatus =
        document.getElementById("playerStatus");

    const playerStatusTop =
        document.getElementById("playerStatusTop");


    /* BACKGROUND */

    const backgroundPicker =
        document.getElementById("backgroundPicker");


    const backgroundPresets =
        document.querySelectorAll(
            ".background-preset"
        );


    /* FORMAT */

    const outputFormatSelect =
        document.getElementById("outputFormatSelect");


    /* =========================================================
       STATE
    ========================================================= */

    let videoDuration = 0;

    let startTime = 0;

    let endTime = 0;

    let activeHandle = null;

    let isDragging = false;

    let isExporting = false;

    let pendingAutoplay = false;

    let processedDownloadUrl = null;


    const MIN_SELECTION = 0.1;


    const ALLOWED_FORMATS = [
        "mp4",
        "mov",
        "webm",
        "avi",
        "mkv",
        "m4v",
        "flv",
        "wmv",
        "gif",
        "mpeg"
    ];


    /*
       PLAYLIST

       Each item contains:

       {
           id,
           file,
           url,
           name,
           duration
       }
    */

    let playlistItems = [];

    let activePlaylistIndex = -1;


    /* =========================================================
       SCREEN MANAGEMENT
    ========================================================= */

    function hideAllScreens() {

        uploadScreen?.classList.add(
            "hidden"
        );

        editorScreen?.classList.add(
            "hidden"
        );

        exportingScreen?.classList.add(
            "hidden"
        );

        resultMode?.classList.add(
            "hidden"
        );
    }


    function showUpload() {

        hideAllScreens();

        uploadScreen?.classList.remove(
            "hidden"
        );
    }


    function showEditor() {

        hideAllScreens();

        editorScreen?.classList.remove(
            "hidden"
        );

        updatePlaybackButtons();
    }


    function showExporting() {

        hideAllScreens();

        exportingScreen?.classList.remove(
            "hidden"
        );
    }


    function showResult() {

        hideAllScreens();

        resultMode?.classList.remove(
            "hidden"
        );
    }


    /* =========================================================
       MESSAGE
    ========================================================= */

    function showMessage(
        message,
        type = "error"
    ) {

        let container =
            document.getElementById(
                "messagesContainer"
            );


        if (!container) {

            container =
                document.createElement(
                    "div"
                );

            container.id =
                "messagesContainer";

            container.className =
                "messages-container";

            document.body.appendChild(
                container
            );
        }


        const item =
            document.createElement(
                "div"
            );

        item.className =
            `django-message ${type}`;

        item.textContent =
            message;


        container.appendChild(
            item
        );


        setTimeout(() => {

            item.style.transition =
                "opacity .3s ease, transform .3s ease";

            item.style.opacity =
                "0";

            item.style.transform =
                "translateY(-8px)";


            setTimeout(() => {

                item.remove();

            }, 350);

        }, 5000);
    }


    document
        .querySelectorAll(
            ".django-message"
        )
        .forEach(
            (message) => {

                setTimeout(() => {

                    message.style.transition =
                        "opacity .3s ease, transform .3s ease";

                    message.style.opacity =
                        "0";

                    message.style.transform =
                        "translateY(-8px)";


                    setTimeout(() => {

                        message.remove();

                    }, 350);

                }, 5000);
            }
        );


    /* =========================================================
       TIME
    ========================================================= */

    function formatTime(seconds) {

        if (
            !Number.isFinite(seconds) ||
            seconds < 0
        ) {
            return "00:00.0";
        }


        const minutes =
            Math.floor(
                seconds / 60
            );


        const remaining =
            seconds -
            minutes * 60;


        return (
            String(minutes).padStart(
                2,
                "0"
            ) +
            ":" +
            remaining
                .toFixed(1)
                .padStart(
                    4,
                    "0"
                )
        );
    }


    function roundTime(value) {

        return Math.round(
            value * 1000
        ) / 1000;
    }


    /* =========================================================
       FORM SYNC
    ========================================================= */

    function syncFormValues() {

        if (startTimeInput) {

            startTimeInput.value =
                roundTime(startTime);
        }


        if (endTimeInput) {

            endTimeInput.value =
                roundTime(endTime);
        }


        if (backgroundInput) {

            backgroundInput.value =
                backgroundPicker?.value ||
                "#000000";
        }


        if (outputFormatInput) {

            outputFormatInput.value =
                outputFormatSelect?.value ||
                "mp4";
        }
    }


    /* =========================================================
       DISPLAY
    ========================================================= */

    function updateDisplays() {

        const selectedDuration =
            Math.max(
                0,
                endTime - startTime
            );


        if (startTimeDisplay) {

            startTimeDisplay.textContent =
                formatTime(startTime);
        }


        if (endTimeDisplay) {

            endTimeDisplay.textContent =
                formatTime(endTime);
        }


        if (durationDisplay) {

            durationDisplay.textContent =
                formatTime(
                    selectedDuration
                );
        }


        if (videoDurationDisplay) {

            videoDurationDisplay.textContent =
                formatTime(
                    videoDuration
                );
        }


        if (timelineStartLabel) {

            timelineStartLabel.textContent =
                formatTime(startTime);
        }


        if (timelineEndLabel) {

            timelineEndLabel.textContent =
                formatTime(endTime);
        }


        if (timelineSelectedValue) {

            timelineSelectedValue.textContent =
                `${formatTime(startTime)} → ${formatTime(endTime)}`;
        }


        syncFormValues();
    }


    /* =========================================================
       TIMELINE
    ========================================================= */

    function timeToPercent(time) {

        if (!videoDuration) {
            return 0;
        }


        return Math.min(
            100,
            Math.max(
                0,
                (time / videoDuration) *
                100
            )
        );
    }


    function percentToTime(percent) {

        if (!videoDuration) {
            return 0;
        }


        return (
            Math.min(
                100,
                Math.max(
                    0,
                    percent
                )
            ) / 100
        ) * videoDuration;
    }


    function renderTimeline() {

        if (!videoDuration) {
            return;
        }


        const startPercent =
            timeToPercent(
                startTime
            );


        const endPercent =
            timeToPercent(
                endTime
            );


        if (startHandle) {

            startHandle.style.left =
                `${startPercent}%`;
        }


        if (endHandle) {

            endHandle.style.left =
                `${endPercent}%`;
        }


        if (timelineSelection) {

            timelineSelection.style.left =
                `${startPercent}%`;

            timelineSelection.style.width =
                `${Math.max(
                    0,
                    endPercent -
                    startPercent
                )}%`;
        }


        updateDisplays();

        updatePlayhead();
    }


    function updatePlayhead() {

        if (
            !sourceVideo ||
            !videoDuration
        ) {
            return;
        }


        const currentTime =
            Number(
                sourceVideo.currentTime
            ) || 0;


        if (timelinePlayhead) {

            timelinePlayhead.style.left =
                `${timeToPercent(
                    currentTime
                )}%`;
        }
    }


    function pointerToPercent(event) {

        if (!timeline) {
            return 0;
        }


        const rect =
            timeline.getBoundingClientRect();


        if (!rect.width) {
            return 0;
        }


        const x =
            event.clientX -
            rect.left;


        return Math.max(
            0,
            Math.min(
                100,
                (x / rect.width) *
                100
            )
        );
    }


    /* =========================================================
       HANDLE DRAGGING
    ========================================================= */

    function startHandleDrag(handle) {

        activeHandle =
            handle;

        isDragging =
            true;

        sourceVideo?.pause();

        document.body.style.userSelect =
            "none";
    }


    function stopHandleDrag() {

        activeHandle =
            null;

        isDragging =
            false;

        document.body.style.userSelect =
            "";
    }


    startHandle?.addEventListener(
        "pointerdown",
        (event) => {

            event.preventDefault();

            startHandle.setPointerCapture?.(
                event.pointerId
            );

            startHandleDrag(
                "start"
            );
        }
    );


    endHandle?.addEventListener(
        "pointerdown",
        (event) => {

            event.preventDefault();

            endHandle.setPointerCapture?.(
                event.pointerId
            );

            startHandleDrag(
                "end"
            );
        }
    );


    document.addEventListener(
        "pointermove",
        (event) => {

            if (
                !isDragging ||
                !activeHandle ||
                !videoDuration
            ) {
                return;
            }


            const percent =
                pointerToPercent(
                    event
                );


            const time =
                percentToTime(
                    percent
                );


            if (
                activeHandle ===
                "start"
            ) {

                startTime =
                    Math.min(
                        roundTime(time),
                        endTime -
                        MIN_SELECTION
                    );


                startTime =
                    Math.max(
                        0,
                        startTime
                    );

            } else {

                endTime =
                    Math.max(
                        roundTime(time),
                        startTime +
                        MIN_SELECTION
                    );


                endTime =
                    Math.min(
                        videoDuration,
                        endTime
                    );
            }


            renderTimeline();


            /*
               When dragging the trim handle,
               show exactly where the trim
               point is in the video.
            */

            sourceVideo.currentTime =
                activeHandle === "start"
                    ? startTime
                    : endTime;
        }
    );


    document.addEventListener(
        "pointerup",
        stopHandleDrag
    );


    document.addEventListener(
        "pointercancel",
        stopHandleDrag
    );


    timeline?.addEventListener(
        "click",
        (event) => {

            if (
                isDragging ||
                event.target ===
                    startHandle ||
                event.target ===
                    endHandle
            ) {
                return;
            }


            if (!videoDuration) {
                return;
            }


            const percent =
                pointerToPercent(
                    event
                );


            const time =
                percentToTime(
                    percent
                );


            /*
               Timeline click moves the video
               playhead but does not change
               trim range.
            */

            sourceVideo.currentTime =
                time;


            updatePlayhead();
        }
    );


    /* =========================================================
       KEYBOARD TRIM CONTROL
    ========================================================= */

    document.addEventListener(
        "keydown",
        (event) => {

            if (
                editorScreen?.classList.contains(
                    "hidden"
                )
            ) {
                return;
            }


            if (!videoDuration) {
                return;
            }


            if (
                event.key !==
                    "ArrowLeft" &&
                event.key !==
                    "ArrowRight"
            ) {
                return;
            }


            const activeElement =
                document.activeElement;


            if (
                activeElement &&
                (
                    activeElement.tagName ===
                        "INPUT" ||
                    activeElement.tagName ===
                        "SELECT" ||
                    activeElement.tagName ===
                        "TEXTAREA"
                )
            ) {
                return;
            }


            event.preventDefault();


            const step =
                0.1;


            const direction =
                event.key ===
                    "ArrowLeft"
                    ? -step
                    : step;


            if (
                activeHandle ===
                "start"
            ) {

                startTime =
                    Math.max(
                        0,
                        Math.min(
                            endTime -
                                MIN_SELECTION,
                            startTime +
                                direction
                        )
                    );

            } else if (
                activeHandle ===
                "end"
            ) {

                endTime =
                    Math.min(
                        videoDuration,
                        Math.max(
                            startTime +
                                MIN_SELECTION,
                            endTime +
                                direction
                        )
                    );

            } else {

                sourceVideo.currentTime =
                    Math.max(
                        startTime,
                        Math.min(
                            endTime,
                            sourceVideo.currentTime +
                                direction
                        )
                    );
            }


            renderTimeline();
        }
    );


    /* =========================================================
       FILE VALIDATION
    ========================================================= */

    function validateVideoFile(file) {

        if (!file) {

            showMessage(
                "Please select a video."
            );

            return false;
        }


        if (
            !file.type ||
            !file.type.startsWith(
                "video/"
            )
        ) {

            showMessage(
                "Please select a valid video file."
            );

            return false;
        }


        if (file.size <= 0) {

            showMessage(
                "The selected video file is empty."
            );

            return false;
        }


        return true;
    }


    /* =========================================================
       PLAYLIST ID
    ========================================================= */

    function createPlaylistId(file) {

        return [
            file.name,
            file.size,
            file.lastModified
        ].join("_");
    }


    /* =========================================================
       PLAYLIST RENDER
    ========================================================= */

    function renderPlaylist() {

        if (!playlist) {
            return;
        }


        playlist.innerHTML =
            "";


        if (playlistCount) {

            playlistCount.textContent =
                playlistItems.length;
        }


        if (!playlistItems.length) {

            const empty =
                document.createElement(
                    "div"
                );

            empty.className =
                "playlist-empty";

            empty.textContent =
                "Add videos to create a playlist.";

            playlist.appendChild(
                empty
            );

            updatePlaybackButtons();

            return;
        }


        playlistItems.forEach(
            (item, index) => {

                const row =
                    document.createElement(
                        "div"
                    );

                row.className =
                    "playlist-item";


                if (
                    index ===
                    activePlaylistIndex
                ) {

                    row.classList.add(
                        "active"
                    );
                }


                const number =
                    document.createElement(
                        "div"
                    );

                number.className =
                    "playlist-number";


                number.textContent =
                    index ===
                        activePlaylistIndex
                        ? "▶"
                        : String(
                            index + 1
                        );


                const details =
                    document.createElement(
                        "div"
                    );

                details.className =
                    "playlist-details";


                const name =
                    document.createElement(
                        "div"
                    );

                name.className =
                    "playlist-name";

                name.textContent =
                    item.name;


                const duration =
                    document.createElement(
                        "div"
                    );

                duration.className =
                    "playlist-duration";

                duration.textContent =
                    item.duration > 0
                        ? formatTime(
                            item.duration
                        )
                        : "Loading...";


                details.appendChild(
                    name
                );

                details.appendChild(
                    duration
                );


                const remove =
                    document.createElement(
                        "button"
                    );

                remove.type =
                    "button";

                remove.className =
                    "playlist-remove";

                remove.title =
                    "Remove video";

                remove.textContent =
                    "×";


                remove.addEventListener(
                    "click",
                    (event) => {

                        event.stopPropagation();

                        removePlaylistItem(
                            index
                        );
                    }
                );


                row.appendChild(
                    number
                );

                row.appendChild(
                    details
                );

                row.appendChild(
                    remove
                );


                row.addEventListener(
                    "click",
                    () => {

                        selectPlaylistItem(
                            index,
                            true
                        );
                    }
                );


                playlist.appendChild(
                    row
                );
            }
        );


        updatePlaybackButtons();
    }


    /* =========================================================
       ADD PLAYLIST FILES
    ========================================================= */

    function addFilesToPlaylist(files) {

        if (
            !files ||
            !files.length
        ) {
            return;
        }


        const newItems = [];


        Array.from(files).forEach(
            (file) => {

                if (
                    !validateVideoFile(
                        file
                    )
                ) {
                    return;
                }


                const id =
                    createPlaylistId(
                        file
                    );


                const existingIndex =
                    playlistItems.findIndex(
                        item =>
                            item.id === id
                    );


                if (
                    existingIndex !==
                    -1
                ) {

                    /*
                       If duplicate file selected,
                       just switch to that playlist item.
                    */

                    selectPlaylistItem(
                        existingIndex,
                        false
                    );

                    return;
                }


                newItems.push({
                    id: id,

                    file: file,

                    url:
                        URL.createObjectURL(
                            file
                        ),

                    name: file.name,

                    duration: 0
                });
            }
        );


        playlistItems.push(
            ...newItems
        );


        renderPlaylist();


        if (
            activePlaylistIndex ===
            -1 &&
            playlistItems.length
        ) {

            selectPlaylistItem(
                0,
                false
            );
        }
    }


    /* =========================================================
       SELECT PLAYLIST ITEM
    ========================================================= */

    function selectPlaylistItem(
        index,
        autoplay = false
    ) {

        if (
            index < 0 ||
            index >=
                playlistItems.length
        ) {
            return;
        }


        const item =
            playlistItems[index];


        activePlaylistIndex =
            index;


        pendingAutoplay =
            Boolean(autoplay);


        /*
           Clear old result URL.
        */

        processedDownloadUrl =
            null;


        if (downloadButton) {

            downloadButton.href =
                "#";

            downloadButton.removeAttribute(
                "download"
            );
        }


        /*
           Stop old video.
        */

        sourceVideo.pause();


        sourceVideo.removeAttribute(
            "src"
        );

        sourceVideo.load();


        /*
           Load new playlist video.
        */

        sourceVideo.src =
            item.url;

        sourceVideo.load();


        /*
           Reset trim range.

           IMPORTANT:
           We do not keep an old video's
           trim range.
        */

        videoDuration =
            0;

        startTime =
            0;

        endTime =
            0;


        if (currentVideoName) {

            currentVideoName.textContent =
                item.name;
        }


        if (videoCurrentName) {

            videoCurrentName.textContent =
                item.name;
        }


        if (videoResolution) {

            videoResolution.textContent =
                "Loading...";
        }


        exportButton.disabled =
            true;


        showEditor();


        renderPlaylist();

        updateDisplays();

        updatePlaybackButtons();
    }


    /* =========================================================
       REMOVE PLAYLIST ITEM
    ========================================================= */

    function removePlaylistItem(index) {

        if (
            index < 0 ||
            index >=
                playlistItems.length
        ) {
            return;
        }


        const wasActive =
            index ===
            activePlaylistIndex;


        const item =
            playlistItems[index];


        if (item.url) {

            URL.revokeObjectURL(
                item.url
            );
        }


        playlistItems.splice(
            index,
            1
        );


        if (!playlistItems.length) {

            activePlaylistIndex =
                -1;

            videoDuration =
                0;

            startTime =
                0;

            endTime =
                0;


            sourceVideo.pause();

            sourceVideo.removeAttribute(
                "src"
            );

            sourceVideo.load();


            exportButton.disabled =
                true;


            renderPlaylist();

            showUpload();

            return;
        }


        if (
            index <
            activePlaylistIndex
        ) {

            activePlaylistIndex--;
        }


        if (wasActive) {

            const nextIndex =
                Math.min(
                    index,
                    playlistItems.length -
                        1
                );


            selectPlaylistItem(
                nextIndex,
                false
            );

            return;
        }


        renderPlaylist();
    }


    /* =========================================================
       UPLOAD BUTTON
    ========================================================= */

    uploadButton?.addEventListener(
        "click",
        () => {

            videoInput?.click();
        }
    );


    videoInput?.addEventListener(
        "change",
        () => {

            const file =
                videoInput.files?.[0];


            if (!file) {
                return;
            }


            addFilesToPlaylist([
                file
            ]);


            /*
               Export uses the playlist file
               directly, so the Django input
               can be cleared safely.
            */

            videoInput.value =
                "";
        }
    );


    /* =========================================================
       PLAYLIST ADD
    ========================================================= */

    addPlaylistButton?.addEventListener(
        "click",
        () => {

            playlistInput?.click();
        }
    );


    playlistInput?.addEventListener(
        "change",
        () => {

            addFilesToPlaylist(
                playlistInput.files
            );


            playlistInput.value =
                "";
        }
    );


    /* =========================================================
       VIDEO METADATA
    ========================================================= */

    sourceVideo?.addEventListener(
        "loadedmetadata",
        () => {

            const duration =
                Number(
                    sourceVideo.duration
                );


            if (
                !Number.isFinite(
                    duration
                ) ||
                duration <= 0
            ) {

                showMessage(
                    "Unable to read the video duration."
                );

                exportButton.disabled =
                    true;

                return;
            }


            videoDuration =
                duration;


            /*
               DEFAULT TRIM RANGE

               Full video.

               Example:

               00:00 → 02:45
            */

            startTime =
                0;

            endTime =
                duration;


            const activeItem =
                playlistItems[
                    activePlaylistIndex
                ];


            if (activeItem) {

                activeItem.duration =
                    duration;
            }


            if (
                sourceVideo.videoWidth &&
                sourceVideo.videoHeight
            ) {

                videoResolution.textContent =
                    `${sourceVideo.videoWidth} × ${sourceVideo.videoHeight}`;
            }


            /*
               Enable export after metadata.
            */

            exportButton.disabled =
                false;


            renderTimeline();

            renderPlaylist();

            updatePlaybackButtons();


            /*
               If playlist item was clicked,
               start playing automatically.
            */

            if (pendingAutoplay) {

                pendingAutoplay =
                    false;


                /*
                   Always start at trim start.
                */

                sourceVideo.currentTime =
                    startTime;


                const playPromise =
                    sourceVideo.play();


                if (
                    playPromise &&
                    typeof playPromise.catch ===
                        "function"
                ) {

                    playPromise.catch(
                        () => {}
                    );
                }
            }
        }
    );


    /* =========================================================
       VIDEO ERROR
    ========================================================= */

    sourceVideo?.addEventListener(
        "error",
        () => {

            showMessage(
                "Unable to load this video."
            );


            exportButton.disabled =
                true;


            updatePlaybackButtons();
        }
    );


    /* =========================================================
       PLAY / PAUSE ICON
    ========================================================= */

    function updatePlayPauseIcon() {

        const playing =
            sourceVideo &&
            !sourceVideo.paused &&
            !sourceVideo.ended;


        const icon =
            playing
                ? "❚❚"
                : "▶";


        if (playPauseButton) {

            playPauseButton.textContent =
                icon;
        }


        if (playPauseButtonTop) {

            playPauseButtonTop.textContent =
                icon;
        }
    }


    /* =========================================================
       PLAYBACK BUTTONS
    ========================================================= */

    function updatePlaybackButtons() {

        const hasVideo =
            activePlaylistIndex >= 0 &&
            activePlaylistIndex <
                playlistItems.length &&
            videoDuration > 0;


        const canPrevious =
            activePlaylistIndex > 0;


        const canNext =
            activePlaylistIndex >= 0 &&
            activePlaylistIndex <
                playlistItems.length -
                1;


        [
            previousVideoButton,
            previousVideoButtonTop
        ].forEach(
            button => {

                if (button) {

                    button.disabled =
                        !canPrevious;
                }
            }
        );


        [
            nextVideoButton,
            nextVideoButtonTop
        ].forEach(
            button => {

                if (button) {

                    button.disabled =
                        !canNext;
                }
            }
        );


        [
            playPauseButton,
            playPauseButtonTop
        ].forEach(
            button => {

                if (button) {

                    button.disabled =
                        !hasVideo;
                }
            }
        );


        updatePlayPauseIcon();
    }


    /* =========================================================
       PLAY SELECTED TRIM RANGE
    ========================================================= */

    function playSelectedRange() {

        if (
            !sourceVideo ||
            !videoDuration
        ) {
            return;
        }


        /*
           If current position is outside
           selected range, start from trim start.
        */

        if (
            sourceVideo.currentTime <
                startTime ||
            sourceVideo.currentTime >=
                endTime
        ) {

            sourceVideo.currentTime =
                startTime;
        }


        const promise =
            sourceVideo.play();


        if (
            promise &&
            typeof promise.catch ===
                "function"
        ) {

            promise.catch(
                () => {}
            );
        }
    }


    function togglePlayPause() {

        if (
            !sourceVideo ||
            !videoDuration
        ) {
            return;
        }


        if (sourceVideo.paused) {

            playSelectedRange();

        } else {

            sourceVideo.pause();
        }
    }


    [
        playPauseButton,
        playPauseButtonTop
    ].forEach(
        button => {

            button?.addEventListener(
                "click",
                togglePlayPause
            );
        }
    );


    /* =========================================================
       PREVIOUS
    ========================================================= */

    function playPreviousVideo() {

        if (
            activePlaylistIndex <= 0
        ) {
            return;
        }


        selectPlaylistItem(
            activePlaylistIndex -
                1,
            true
        );
    }


    [
        previousVideoButton,
        previousVideoButtonTop
    ].forEach(
        button => {

            button?.addEventListener(
                "click",
                playPreviousVideo
            );
        }
    );


    /* =========================================================
       NEXT
    ========================================================= */

    function playNextVideo() {

        if (
            activePlaylistIndex < 0 ||
            activePlaylistIndex >=
                playlistItems.length -
                1
        ) {
            return;
        }


        selectPlaylistItem(
            activePlaylistIndex +
                1,
            true
        );
    }


    [
        nextVideoButton,
        nextVideoButtonTop
    ].forEach(
        button => {

            button?.addEventListener(
                "click",
                playNextVideo
            );
        }
    );


    /* =========================================================
       VIDEO PLAY
    ========================================================= */

    sourceVideo?.addEventListener(
        "play",
        () => {

            /*
               Always enforce selected trim range.
            */

            if (
                sourceVideo.currentTime <
                    startTime ||
                sourceVideo.currentTime >=
                    endTime
            ) {

                sourceVideo.currentTime =
                    startTime;
            }


            updatePlayPauseIcon();

            setPlayerStatus(
                "Playing selected range"
            );
        }
    );


    /* =========================================================
       VIDEO PAUSE
    ========================================================= */

    sourceVideo?.addEventListener(
        "pause",
        () => {

            updatePlayPauseIcon();

            setPlayerStatus(
                "Paused"
            );
        }
    );


    /* =========================================================
       VIDEO TIMEUPDATE
    ========================================================= */

    sourceVideo?.addEventListener(
        "timeupdate",
        () => {

            updatePlayhead();


            /*
               THIS IS THE IMPORTANT TRIM
               PREVIEW BEHAVIOUR.

               Video cannot play beyond
               the selected END handle.
            */

            if (
                !sourceVideo.paused &&
                videoDuration > 0 &&
                endTime > startTime &&
                sourceVideo.currentTime >=
                    endTime
            ) {

                /*
                   Loop back to trim START.

                   This makes it very easy for
                   the user to preview the exact
                   selected range repeatedly.
                */

                sourceVideo.currentTime =
                    startTime;


                updatePlayhead();
            }
        }
    );


    /* =========================================================
       VIDEO ENDED
    ========================================================= */

    sourceVideo?.addEventListener(
        "ended",
        () => {

            /*
               If the entire video is selected,
               move to the next playlist item.
            */

            if (
                activePlaylistIndex >= 0 &&
                activePlaylistIndex <
                    playlistItems.length -
                    1 &&
                Math.abs(
                    endTime -
                    videoDuration
                ) < 0.05
            ) {

                selectPlaylistItem(
                    activePlaylistIndex +
                        1,
                    true
                );

                return;
            }


            sourceVideo.currentTime =
                startTime;


            updatePlayPauseIcon();

            updatePlayhead();
        }
    );


    /* =========================================================
       PLAYER STATUS
    ========================================================= */

    function setPlayerStatus(text) {

        if (playerStatus) {

            playerStatus.textContent =
                text;
        }


        if (playerStatusTop) {

            playerStatusTop.textContent =
                text;
        }
    }


    /* =========================================================
       BACKGROUND
    ========================================================= */

    function isValidHexColor(value) {

        return /^#[0-9A-F]{6}$/i.test(
            value
        );
    }


    function setBackgroundColor(color) {

        if (
            !isValidHexColor(
                color
            )
        ) {
            return;
        }


        if (backgroundPicker) {

            backgroundPicker.value =
                color;
        }


        if (backgroundInput) {

            backgroundInput.value =
                color;
        }


        const videoPanel =
            document.getElementById(
                "editorVideoPanel"
            );


        if (videoPanel) {

            videoPanel.style.backgroundColor =
                color;
        }


        backgroundPresets.forEach(
            preset => {

                preset.classList.toggle(
                    "active",
                    preset.dataset.color
                        ?.toLowerCase() ===
                    color.toLowerCase()
                );
            }
        );
    }


    backgroundPresets.forEach(
        preset => {

            preset.addEventListener(
                "click",
                () => {

                    const color =
                        preset.dataset.color;

                    if (color) {

                        setBackgroundColor(
                            color
                        );
                    }
                }
            );
        }
    );


    backgroundPicker?.addEventListener(
        "input",
        () => {

            setBackgroundColor(
                backgroundPicker.value
            );
        }
    );


    /* =========================================================
       FORMAT
    ========================================================= */

    function setOutputFormat(format) {

        format =
            String(
                format || ""
            )
                .toLowerCase()
                .trim();


        if (
            !ALLOWED_FORMATS.includes(
                format
            )
        ) {

            format =
                "mp4";
        }


        if (outputFormatSelect) {

            outputFormatSelect.value =
                format;
        }


        if (outputFormatInput) {

            outputFormatInput.value =
                format;
        }
    }


    outputFormatSelect?.addEventListener(
        "change",
        () => {

            setOutputFormat(
                outputFormatSelect.value
            );
        }
    );


    /* =========================================================
       EXPORT VALIDATION
    ========================================================= */

    function validateExport() {

        const activeItem =
            playlistItems[
                activePlaylistIndex
            ];


        if (!activeItem?.file) {

            showMessage(
                "Please select a video first."
            );

            return false;
        }


        if (
            !validateVideoFile(
                activeItem.file
            )
        ) {
            return false;
        }


        if (!videoDuration) {

            showMessage(
                "Video is still loading."
            );

            return false;
        }


        if (
            startTime < 0 ||
            endTime <= startTime
        ) {

            showMessage(
                "Please select a valid trim range."
            );

            return false;
        }


        if (
            endTime -
                startTime <
            MIN_SELECTION
        ) {

            showMessage(
                "The selected range is too short."
            );

            return false;
        }


        if (
            endTime >
            videoDuration +
                0.001
        ) {

            showMessage(
                "The end time is invalid."
            );

            return false;
        }


        const format =
            outputFormatSelect?.value ||
            "mp4";


        if (
            !ALLOWED_FORMATS.includes(
                format
            )
        ) {

            showMessage(
                "Please select a valid output format."
            );

            return false;
        }


        const background =
            backgroundPicker?.value ||
            "#000000";


        if (
            !isValidHexColor(
                background
            )
        ) {

            showMessage(
                "Please select a valid background color."
            );

            return false;
        }


        return true;
    }


    /* =========================================================
       RESULT VALUES
    ========================================================= */

    function updateResultValues() {

        const format =
            outputFormatSelect?.value ||
            "mp4";


        if (resultStartDisplay) {

            resultStartDisplay.textContent =
                formatTime(
                    startTime
                );
        }


        if (resultEndDisplay) {

            resultEndDisplay.textContent =
                formatTime(
                    endTime
                );
        }


        if (resultDurationDisplay) {

            resultDurationDisplay.textContent =
                formatTime(
                    endTime -
                    startTime
                );
        }


        if (resultFormatDisplay) {

            resultFormatDisplay.textContent =
                format.toUpperCase();
        }
    }


    /* =========================================================
       CSRF
    ========================================================= */

    function getCSRFToken() {

        return trimForm
            ?.querySelector(
                'input[name="csrfmiddlewaretoken"]'
            )
            ?.value || "";
    }


    /* =========================================================
       EXPORT
    ========================================================= */

    exportButton?.addEventListener(
        "click",
        async () => {

            if (isExporting) {
                return;
            }


            if (!validateExport()) {
                return;
            }


            const activeItem =
                playlistItems[
                    activePlaylistIndex
                ];


            if (!activeItem?.file) {
                return;
            }


            isExporting =
                true;


            exportButton.disabled =
                true;


            /*
               Clear previous download.

               Old URL must NEVER remain available
               during the new processing request.
            */

            processedDownloadUrl =
                null;


            if (downloadButton) {

                downloadButton.href =
                    "#";

                downloadButton.removeAttribute(
                    "download"
                );
            }


            sourceVideo.pause();


            syncFormValues();


            showExporting();


            try {

                /*
                   Use existing Django form.
                */

                const formData =
                    new FormData(
                        trimForm
                    );


                /*
                   Replace old video with
                   currently selected playlist video.
                */

                formData.delete(
                    "video"
                );


                formData.set(
                    "video",
                    activeItem.file,
                    activeItem.file.name
                );


                formData.set(
                    "start_time",
                    String(
                        roundTime(
                            startTime
                        )
                    )
                );


                formData.set(
                    "end_time",
                    String(
                        roundTime(
                            endTime
                        )
                    )
                );


                formData.set(
                    "background_color",
                    backgroundPicker?.value ||
                        "#000000"
                );


                formData.set(
                    "output_format",
                    outputFormatSelect?.value ||
                        "mp4"
                );


                const csrfToken =
                    getCSRFToken();


                if (!csrfToken) {

                    throw new Error(
                        "CSRF token is missing."
                    );
                }


                const response =
                    await fetch(
                        trimForm.action,
                        {
                            method: "POST",

                            body: formData,

                            credentials:
                                "same-origin",

                            headers: {

                                "X-CSRFToken":
                                    csrfToken,

                                "X-Requested-With":
                                    "XMLHttpRequest"
                            }
                        }
                    );


                let data;


                try {

                    data =
                        await response.json();

                } catch {

                    throw new Error(
                        "Invalid response from server."
                    );
                }


                if (
                    !response.ok ||
                    !data ||
                    data.success !== true
                ) {

                    throw new Error(
                        data?.error ||
                        data?.message ||
                        "Video processing failed."
                    );
                }


                if (!data.video_url) {

                    throw new Error(
                        "Processed video URL was not returned."
                    );
                }


                /*
                   Fresh preview URL.
                */

                const separator =
                    data.video_url.includes(
                        "?"
                    )
                        ? "&"
                        : "?";


                const freshVideoUrl =
                    `${data.video_url}${separator}trim_preview=${Date.now()}`;


                resultVideo.pause();


                resultVideo.removeAttribute(
                    "src"
                );


                resultVideo.load();


                resultVideo.src =
                    freshVideoUrl;


                resultVideo.load();


                /*
                   Set download ONLY after
                   processing succeeds.
                */

                const currentDownloadUrl =
                    data.download_url ||
                    data.video_url;


                processedDownloadUrl =
                    currentDownloadUrl;


                if (downloadButton) {

                    downloadButton.href =
                        currentDownloadUrl;


                    if (data.filename) {

                        downloadButton.download =
                            data.filename;

                    } else {

                        downloadButton.removeAttribute(
                            "download"
                        );
                    }
                }


                updateResultValues();


                showResult();


                resultVideo.addEventListener(
                    "loadedmetadata",
                    () => {

                        try {

                            resultVideo.currentTime =
                                0;

                        } catch {
                            // Ignore.
                        }

                    },
                    {
                        once: true
                    }
                );

            } catch (error) {

                console.error(
                    "Trim export error:",
                    error
                );


                showEditor();


                showMessage(
                    error?.message ||
                    "Video processing failed."
                );

            } finally {

                isExporting =
                    false;


                exportButton.disabled =
                    !videoDuration;


                updatePlaybackButtons();
            }
        }
    );


    /* =========================================================
       DOWNLOAD SAFETY
    ========================================================= */

    downloadButton?.addEventListener(
        "click",
        (event) => {

            if (!processedDownloadUrl) {

                event.preventDefault();


                showMessage(
                    "Your video is still being processed."
                );
            }
        }
    );


    /* =========================================================
       EDIT AGAIN
    ========================================================= */

    editAgainButton?.addEventListener(
        "click",
        () => {

            processedDownloadUrl =
                null;


            if (downloadButton) {

                downloadButton.href =
                    "#";

                downloadButton.removeAttribute(
                    "download"
                );
            }


            showEditor();


            renderTimeline();


            if (
                sourceVideo &&
                videoDuration
            ) {

                sourceVideo.currentTime =
                    startTime;
            }


            setPlayerStatus(
                "Ready"
            );


            updatePlaybackButtons();
        }
    );


    /* =========================================================
       INITIAL STATE
    ========================================================= */

    exportButton.disabled =
        true;


    setBackgroundColor(
        backgroundPicker?.value ||
        "#000000"
    );


    setOutputFormat(
        outputFormatSelect?.value ||
        "mp4"
    );


    updateDisplays();

    renderPlaylist();

    updatePlaybackButtons();


    /* =========================================================
       CLEANUP
    ========================================================= */

    window.addEventListener(
        "beforeunload",
        () => {

            playlistItems.forEach(
                item => {

                    if (item.url) {

                        URL.revokeObjectURL(
                            item.url
                        );
                    }
                }
            );
        }
    );

});