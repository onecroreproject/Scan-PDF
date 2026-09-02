document.addEventListener("DOMContentLoaded", function () {
    "use strict";

    /* =========================================
       ELEMENTS
    ========================================= */

    const form = document.getElementById("resizeForm");

    const videoInput = document.getElementById("id_video");

    const uploadSection =
        document.getElementById("uploadSection");

    const editorSection =
        document.getElementById("editorSection");

    const processingSection =
        document.getElementById("processingSection");

    const resultSection =
        document.getElementById("resultSection");

    const videoPreview =
        document.getElementById("videoPreview");

    const videoStage =
        document.getElementById("videoStage");

    const editorFileName =
        document.getElementById("editorFileName");

    const sourceDimensions =
        document.getElementById("sourceDimensions");

    const currentDimensions =
        document.getElementById("currentDimensions");

    const widthControl =
        document.getElementById("widthControl");

    const heightControl =
        document.getElementById("heightControl");

    const lockAspectRatio =
        document.getElementById("lockAspectRatio");

    const zoomMinus =
        document.getElementById("zoomMinus");

    const zoomPlus =
        document.getElementById("zoomPlus");

    const zoomValue =
        document.getElementById("zoomValue");

    const changeVideoButton =
        document.getElementById("changeVideoButton");

    const resizeButton =
        document.getElementById("resizeButton");


    /* =========================================
       DJANGO HIDDEN FIELDS
    ========================================= */

    const hiddenWidth =
        document.getElementById("id_width");

    const hiddenHeight =
        document.getElementById("id_height");

    const hiddenAspectRatio =
        document.getElementById("id_aspect_ratio");

    const hiddenFitMode =
        document.getElementById("id_fit_mode");

    const hiddenZoom =
        document.getElementById("id_zoom");

    const hiddenPositionX =
        document.getElementById("id_position_x");

    const hiddenPositionY =
        document.getElementById("id_position_y");

    const hiddenBackgroundColor =
        document.getElementById(
            "id_background_color"
        );

    const hiddenOutputFormat =
        document.getElementById(
            "id_output_format"
        );


    /* =========================================
       EDITOR STATE
    ========================================= */

    let state = {
        width: 1280,
        height: 720,

        aspectRatio: "16:9",

        fitMode: "fit",

        zoom: 1,

        positionX: 0,
        positionY: 0,

        backgroundColor: "#000000",

        outputFormat: "mp4",

        sourceWidth: 0,
        sourceHeight: 0,

        objectUrl: null
    };


    /* =========================================
       UTILITY
    ========================================= */

    function clamp(
        value,
        min,
        max
    ) {
        return Math.min(
            Math.max(value, min),
            max
        );
    }


    function parseRatio(
        ratio
    ) {
        try {

            const parts =
                ratio.split(":");

            if (parts.length !== 2) {
                return null;
            }

            const width =
                Number(parts[0]);

            const height =
                Number(parts[1]);

            if (
                width <= 0 ||
                height <= 0
            ) {
                return null;
            }

            return width / height;

        } catch (error) {

            console.error(
                "Unable to parse aspect ratio:",
                error
            );

            return null;
        }
    }


    function updateHiddenFields() {

        if (hiddenWidth) {
            hiddenWidth.value =
                state.width;
        }

        if (hiddenHeight) {
            hiddenHeight.value =
                state.height;
        }

        if (hiddenAspectRatio) {
            hiddenAspectRatio.value =
                state.aspectRatio;
        }

        if (hiddenFitMode) {
            hiddenFitMode.value =
                state.fitMode;
        }

        if (hiddenZoom) {
            hiddenZoom.value =
                state.zoom.toFixed(2);
        }

        if (hiddenPositionX) {
            hiddenPositionX.value =
                Math.round(state.positionX);
        }

        if (hiddenPositionY) {
            hiddenPositionY.value =
                Math.round(state.positionY);
        }

        if (hiddenBackgroundColor) {
            hiddenBackgroundColor.value =
                state.backgroundColor;
        }

        if (hiddenOutputFormat) {
            hiddenOutputFormat.value =
                state.outputFormat;
        }
    }


    function updateDimensionDisplay() {

        if (currentDimensions) {

            currentDimensions.textContent =
                `Output: ${state.width} × ${state.height}`;

        }
    }


    /* =========================================
       VIDEO UPLOAD
    ========================================= */

    if (videoInput) {

        videoInput.addEventListener(
            "change",
            function () {
 resetPreviousResult();
                try {

                    const file =
                        videoInput.files[0];

                    if (!file) {
                        return;
                    }

                    if (
                        !file.type.startsWith(
                            "video/"
                        )
                    ) {

                        alert(
                            "Please select a valid video file."
                        );

                        videoInput.value = "";

                        return;
                    }


                    /* Revoke previous object URL */

                    if (state.objectUrl) {

                        URL.revokeObjectURL(
                            state.objectUrl
                        );

                    }


                    state.objectUrl =
                        URL.createObjectURL(
                            file
                        );

                    videoPreview.src =
                        state.objectUrl;

                    editorFileName.textContent =
                        file.name;


                    videoPreview.load();


                    /*
                     * Wait until browser has
                     * loaded video metadata.
                     */

                    videoPreview.addEventListener(
                        "loadedmetadata",
                        handleVideoMetadata,
                        {
                            once: true
                        }
                    );


                    /*
                     * Switch upload UI
                     * to editor UI.
                     */

                    uploadSection.style.display =
                        "none";

                    editorSection.style.display =
                        "block";


                } catch (error) {

                    console.error(
                        "Video upload error:",
                        error
                    );

                    alert(
                        "Unable to load the selected video."
                    );

                }

            }
        );

    }


    /* =========================================
       VIDEO METADATA
    ========================================= */

    function handleVideoMetadata() {

        try {

            const width =
                videoPreview.videoWidth;

            const height =
                videoPreview.videoHeight;


            if (
                !width ||
                !height
            ) {

                throw new Error(
                    "Invalid video dimensions."
                );

            }


            state.sourceWidth =
                width;

            state.sourceHeight =
                height;


            sourceDimensions.textContent =
                `Original: ${width} × ${height}`;


            /*
             * Use original dimensions
             * as initial dimensions.
             */

            state.width = width;
            state.height = height;


            widthControl.value =
                width;

            heightControl.value =
                height;


            /*
             * Detect initial ratio.
             */

            state.aspectRatio =
                findClosestAspectRatio(
                    width / height
                );


            setActiveRatio(
                state.aspectRatio
            );


            updateStage();

            updateHiddenFields();

        } catch (error) {

            console.error(
                "Video metadata error:",
                error
            );

            alert(
                "Unable to read video information."
            );

        }

    }

    function resetPreviousResult() {
    const resultSection = document.getElementById("resultSection");

    if (resultSection) {
        resultSection.style.display = "none";
    }

    const downloadButton = resultSection
        ? resultSection.querySelector("a")
        : null;

    if (downloadButton) {
        downloadButton.removeAttribute("href");
    }
}


    /* =========================================
       ASPECT RATIO
    ========================================= */

    const ratioButtons =
        document.querySelectorAll(
            ".ratio-button"
        );


    ratioButtons.forEach(
        function (button) {

            button.addEventListener(
                "click",
                function () {

                    try {

                        const ratio =
                            button.dataset.ratio;

                        if (!ratio) {
                            return;
                        }

                        state.aspectRatio =
                            ratio;

                        const ratioValue =
                            parseRatio(
                                ratio
                            );

                        if (
                            ratioValue === null
                        ) {
                            return;
                        }


                        /*
                         * Change height based
                         * on current width.
                         */

                        state.height =
                            Math.max(
                                2,
                                Math.round(
                                    state.width /
                                    ratioValue
                                )
                            );


                        heightControl.value =
                            state.height;


                        setActiveRatio(
                            ratio
                        );


                        updateStage();

                        updateHiddenFields();

                    } catch (error) {

                        console.error(
                            "Aspect ratio error:",
                            error
                        );

                    }

                }
            );

        }
    );


    function setActiveRatio(
        ratio
    ) {

        ratioButtons.forEach(
            function (button) {

                button.classList.toggle(
                    "active",
                    button.dataset.ratio ===
                    ratio
                );

            }
        );

    }


    function findClosestAspectRatio(
        value
    ) {

        const ratios = [
            {
                name: "16:9",
                value: 16 / 9
            },
            {
                name: "3:2",
                value: 3 / 2
            },
            {
                name: "4:3",
                value: 4 / 3
            },
            {
                name: "1:1",
                value: 1
            },
            {
                name: "4:5",
                value: 4 / 5
            },
            {
                name: "9:16",
                value: 9 / 16
            },
            {
                name: "21:9",
                value: 21 / 9
            }
        ];


        let closest =
            ratios[0];

        let difference =
            Math.abs(
                value -
                closest.value
            );


        ratios.forEach(
            function (ratio) {

                const currentDifference =
                    Math.abs(
                        value -
                        ratio.value
                    );

                if (
                    currentDifference <
                    difference
                ) {

                    closest =
                        ratio;

                    difference =
                        currentDifference;

                }

            }
        );


        return closest.name;
    }


    /* =========================================
       WIDTH
    ========================================= */

    if (widthControl) {

        widthControl.addEventListener(
            "input",
            function () {

                let width =
                    parseInt(
                        widthControl.value,
                        10
                    );


                if (
                    Number.isNaN(width) ||
                    width < 2
                ) {
                    return;
                }


                const oldWidth =
                    state.width;

                state.width =
                    width;


                /*
                 * Keep aspect ratio locked.
                 */

                if (
                    lockAspectRatio.checked
                ) {

                    const ratio =
                        parseRatio(
                            state.aspectRatio
                        );


                    if (ratio) {

                        state.height =
                            Math.max(
                                2,
                                Math.round(
                                    width /
                                    ratio
                                )
                            );

                        heightControl.value =
                            state.height;

                    }

                }


                updateStage();

                updateHiddenFields();

            }
        );

    }


    /* =========================================
       HEIGHT
    ========================================= */

    if (heightControl) {

        heightControl.addEventListener(
            "input",
            function () {

                let height =
                    parseInt(
                        heightControl.value,
                        10
                    );


                if (
                    Number.isNaN(height) ||
                    height < 2
                ) {
                    return;
                }


                state.height =
                    height;


                /*
                 * Keep aspect ratio locked.
                 */

                if (
                    lockAspectRatio.checked
                ) {

                    const ratio =
                        parseRatio(
                            state.aspectRatio
                        );


                    if (ratio) {

                        state.width =
                            Math.max(
                                2,
                                Math.round(
                                    height *
                                    ratio
                                )
                            );

                        widthControl.value =
                            state.width;

                    }

                }


                updateStage();

                updateHiddenFields();

            }
        );

    }


    /* =========================================
       FIT / FILL
    ========================================= */

    const fitButtons =
        document.querySelectorAll(
            ".segment-button"
        );


    fitButtons.forEach(
        function (button) {

            button.addEventListener(
                "click",
                function () {

                    try {

                        const mode =
                            button.dataset.fit;

                        if (!mode) {
                            return;
                        }

                        state.fitMode =
                            mode;


                        fitButtons.forEach(
                            function (item) {

                                item.classList.toggle(
                                    "active",
                                    item === button
                                );

                            }
                        );


                        updateStage();

                        updateHiddenFields();

                    } catch (error) {

                        console.error(
                            "Fit mode error:",
                            error
                        );

                    }

                }
            );

        }
    );


    /* =========================================
       ZOOM
    ========================================= */

    function updateZoomDisplay() {

        const percentage =
            Math.round(
                state.zoom * 100
            );

        zoomValue.textContent =
            `${percentage}%`;

    }


    if (zoomMinus) {

        zoomMinus.addEventListener(
            "click",
            function () {

                state.zoom =
                    clamp(
                        state.zoom - 0.1,
                        0.1,
                        5
                    );

                updateZoomDisplay();

                updateStage();

                updateHiddenFields();

            }
        );

    }


    if (zoomPlus) {

        zoomPlus.addEventListener(
            "click",
            function () {

                state.zoom =
                    clamp(
                        state.zoom + 0.1,
                        0.1,
                        5
                    );

                updateZoomDisplay();

                updateStage();

                updateHiddenFields();

            }
        );

    }


    /* =========================================
       POSITION
    ========================================= */

    const positionButtons =
        document.querySelectorAll(
            ".position-button"
        );


    positionButtons.forEach(
        function (button) {

            button.addEventListener(
                "click",
                function () {

                    try {

                        const position =
                            button.dataset.position;


                        const movement =
                            20;


                        switch (position) {

                            case "left":

                                state.positionX -=
                                    movement;

                                break;


                            case "right":

                                state.positionX +=
                                    movement;

                                break;


                            case "up":

                                state.positionY -=
                                    movement;

                                break;


                            case "down":

                                state.positionY +=
                                    movement;

                                break;


                            case "center":

                                state.positionX =
                                    0;

                                state.positionY =
                                    0;

                                break;


                            default:

                                return;

                        }


                        updateStage();

                        updateHiddenFields();

                    } catch (error) {

                        console.error(
                            "Position error:",
                            error
                        );

                    }

                }
            );

        }
    );


    /* =========================================
       BACKGROUND COLOR
    ========================================= */

    const colorButtons =
        document.querySelectorAll(
            ".color-button"
        );


    colorButtons.forEach(
        function (button) {

            button.addEventListener(
                "click",
                function () {

                    try {

                        const color =
                            button.dataset.color;

                        if (!color) {
                            return;
                        }


                        state.backgroundColor =
                            color;


                        colorButtons.forEach(
                            function (item) {

                                item.classList.toggle(
                                    "active",
                                    item === button
                                );

                            }
                        );


                        videoStage.style.background =
                            state.backgroundColor;


                        updateHiddenFields();

                    } catch (error) {

                        console.error(
                            "Background color error:",
                            error
                        );

                    }

                }
            );

        }
    );


    /* =========================================
       OUTPUT FORMAT
    ========================================= */

    const formatButtons =
        document.querySelectorAll(
            ".format-button"
        );


    formatButtons.forEach(
        function (button) {

            button.addEventListener(
                "click",
                function () {

                    try {

                        const format =
                            button.dataset.format;

                        if (!format) {
                            return;
                        }


                        state.outputFormat =
                            format;


                        formatButtons.forEach(
                            function (item) {

                                item.classList.toggle(
                                    "active",
                                    item === button
                                );

                            }
                        );


                        updateHiddenFields();

                    } catch (error) {

                        console.error(
                            "Output format error:",
                            error
                        );

                    }

                }
            );

        }
    );


    /* =========================================
       REALTIME VIDEO PREVIEW
    ========================================= */

    function updateStage() {

        try {

            /*
             * Keep the editor stage inside
             * the available workspace.
             */

            const wrapper =
                document.querySelector(
                    ".video-stage-wrapper"
                );


            if (!wrapper) {
                return;
            }


            const wrapperWidth =
                wrapper.clientWidth;

            const wrapperHeight =
                wrapper.clientHeight;


            if (
                wrapperWidth <= 0 ||
                wrapperHeight <= 0
            ) {
                return;
            }


            const ratio =
                state.width /
                state.height;


            let stageWidth =
                wrapperWidth * 0.85;

            let stageHeight =
                stageWidth / ratio;


            if (
                stageHeight >
                wrapperHeight * 0.85
            ) {

                stageHeight =
                    wrapperHeight * 0.85;

                stageWidth =
                    stageHeight * ratio;

            }


            videoStage.style.width =
                `${Math.max(
                    20,
                    stageWidth
                )}px`;

            videoStage.style.height =
                `${Math.max(
                    20,
                    stageHeight
                )}px`;


            videoStage.style.background =
                state.backgroundColor;


            /*
             * Preview Fit / Fill.
             */

            if (
                state.fitMode === "fill"
            ) {

                videoPreview.style.width =
                    "100%";

                videoPreview.style.height =
                    "100%";

                videoPreview.style.objectFit =
                    "cover";

            } else {

                videoPreview.style.width =
                    "100%";

                videoPreview.style.height =
                    "100%";

                videoPreview.style.objectFit =
                    "contain";

            }


            /*
             * Apply realtime zoom
             * and position.
             */

            videoPreview.style.transform =
                `translate(
                    ${state.positionX}px,
                    ${state.positionY}px
                )
                scale(${state.zoom})`;


            updateDimensionDisplay();

        } catch (error) {

            console.error(
                "Preview update error:",
                error
            );

        }

    }


    /* =========================================
       CHANGE VIDEO
    ========================================= */

    if (changeVideoButton) {

        changeVideoButton.addEventListener(
            "click",
            function () {

                try {

                    if (state.objectUrl) {

                        URL.revokeObjectURL(
                            state.objectUrl
                        );

                        state.objectUrl =
                            null;

                    }


                    videoPreview.pause();

                    videoPreview.removeAttribute(
                        "src"
                    );

                    videoPreview.load();

                    videoInput.value = "";


                    editorSection.style.display =
                        "none";

                    uploadSection.style.display =
                        "block";


                } catch (error) {

                    console.error(
                        "Change video error:",
                        error
                    );

                }

            }
        );

    }


    /* =========================================
       FORM SUBMISSION
    ========================================= */

    if (form) {

        form.addEventListener(
            "submit",
            function (event) {

                try {
 resetPreviousResult();
                    /*
                     * Make sure final state is
                     * copied to Django fields.
                     */

                    updateHiddenFields();


                    /*
                     * Basic frontend validation.
                     * Django performs the final
                     * validation on the server.
                     */

                    if (
                        !videoInput.files.length
                    ) {

                        event.preventDefault();

                        alert(
                            "Please select a video."
                        );

                        return;

                    }


                    if (
                        state.width < 2 ||
                        state.height < 2
                    ) {

                        event.preventDefault();

                        alert(
                            "Width and height must be at least 2 pixels."
                        );

                        return;

                    }


                    if (
                        state.zoom < 0.1 ||
                        state.zoom > 5
                    ) {

                        event.preventDefault();

                        alert(
                            "Zoom must be between 10% and 500%."
                        );

                        return;

                    }


                    /*
                     * Show processing UI.
                     */

                    editorSection.style.display =
                        "none";

                    processingSection.style.display =
                        "block";


                    /*
                     * Prevent double-click /
                     * duplicate submission.
                     */

                    resizeButton.disabled =
                        true;

                    resizeButton.textContent =
                        "Processing...";


                } catch (error) {

                    console.error(
                        "Form submission error:",
                        error
                    );

                    event.preventDefault();

                    processingSection.style.display =
                        "none";

                    editorSection.style.display =
                        "block";

                }

            }
        );

    }


    /* =========================================
       INITIAL STATE
    ========================================= */

    updateHiddenFields();

    updateZoomDisplay();

});