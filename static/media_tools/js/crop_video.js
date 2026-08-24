(function () {
    "use strict";

    /*
    ============================================================
    Crop Video Editor
    ============================================================

    Responsibilities:
        - Video preview
        - Crop rectangle
        - Dragging
        - Resizing
        - Aspect-ratio handling
        - Custom dimensions
        - Rotation preview
        - Hidden form values
        - Processing spinner

    The actual video processing is done by Django/PyAV.
    ============================================================
    */

    document.addEventListener("DOMContentLoaded", function () {

        try {
            initCropEditor();
        } catch (error) {
            console.error(
                "Crop editor initialization failed:",
                error
            );
        }

    });


    function initCropEditor() {

        /*
        ========================================================
        DOM elements
        ========================================================
        */

        const form = document.getElementById(
            "cropForm"
        );

        const videoInput = document.getElementById(
            "id_video"
        );

        const editorSection = document.getElementById(
            "editorSection"
        );

        const videoEditor = document.getElementById(
            "videoEditor"
        );

        const videoPreview = document.getElementById(
            "videoPreview"
        );

        const cropBox = document.getElementById(
            "cropBox"
        );

        const originalDimensions =
            document.getElementById(
                "originalDimensions"
            );

        const cropWidthDisplay =
            document.getElementById(
                "cropWidthDisplay"
            );

        const cropHeightDisplay =
            document.getElementById(
                "cropHeightDisplay"
            );

        const cropXDisplay =
            document.getElementById(
                "cropXDisplay"
            );

        const cropYDisplay =
            document.getElementById(
                "cropYDisplay"
            );

        const cropRatioInput =
            document.getElementById(
                "cropRatioInput"
            );

        const cropXInput =
            document.getElementById(
                "cropXInput"
            );

        const cropYInput =
            document.getElementById(
                "cropYInput"
            );

        const cropWidthInput =
            document.getElementById(
                "cropWidthInput"
            );

        const cropHeightInput =
            document.getElementById(
                "cropHeightInput"
            );

        const rotationInput =
            document.getElementById(
                "rotationInput"
            );

        const rotationDisplay =
            document.getElementById(
                "rotationDisplay"
            );

        const customSection =
            document.getElementById(
                "customSection"
            );

        const customWidthInput =
            document.getElementById(
                "id_custom_width"
            );

        const customHeightInput =
            document.getElementById(
                "id_custom_height"
            );

        const rotateLeft =
            document.getElementById(
                "rotateLeft"
            );

        const rotateRight =
            document.getElementById(
                "rotateRight"
            );

        const cropSubmit =
            document.getElementById(
                "cropSubmit"
            );

        const processingOverlay =
            document.getElementById(
                "processingOverlay"
            );

        const ratioButtons =
            document.querySelectorAll(
                ".ratio-btn"
            );


        /*
        ========================================================
        Validate required elements
        ========================================================
        */

        if (!form) {
            throw new Error(
                "Crop form was not found."
            );
        }

        if (!videoInput) {
            throw new Error(
                "Video input was not found."
            );
        }

        if (!videoPreview) {
            throw new Error(
                "Video preview was not found."
            );
        }

        if (!cropBox) {
            throw new Error(
                "Crop box was not found."
            );
        }


        /*
        ========================================================
        Editor state
        ========================================================
        */

        let videoWidth = 0;
        let videoHeight = 0;

        let crop = {
            x: 0,
            y: 0,
            width: 0,
            height: 0
        };

        let activeRatio = "free";

        let rotation = 0;

        let objectUrl = null;

        let interaction = null;


        /*
        ========================================================
        Ratio definitions
        ========================================================
        */

        const RATIOS = {
            "1:1": {
                width: 1,
                height: 1
            },

            "2:2": {
                width: 2,
                height: 2
            },

            "3:7": {
                width: 3,
                height: 7
            },

            "8:5": {
                width: 8,
                height: 5
            },

            "4:3": {
                width: 4,
                height: 3
            },

            "16:9": {
                width: 16,
                height: 9
            },

            "9:16": {
                width: 9,
                height: 16
            }
        };


        /*
        ========================================================
        Video upload
        ========================================================
        */

        videoInput.addEventListener(
            "change",
            function () {

                try {

                    const file =
                        videoInput.files &&
                        videoInput.files[0];

                    if (!file) {
                        resetEditor();
                        return;
                    }

                    validateVideoFile(file);

                    if (objectUrl) {
                        URL.revokeObjectURL(
                            objectUrl
                        );
                    }

                    objectUrl =
                        URL.createObjectURL(
                            file
                        );

                    videoPreview.src =
                        objectUrl;

                    videoPreview.load();

                    editorSection.style.display =
                        "block";

                } catch (error) {

                    console.error(
                        "Video selection error:",
                        error
                    );

                    alert(
                        error.message ||
                        "Unable to load the selected video."
                    );

                    resetEditor();
                }

            }
        );


        /*
        ========================================================
        Validate uploaded file
        ========================================================
        */

        function validateVideoFile(file) {

            const maxSize =
                500 * 1024 * 1024;

            const allowedExtensions = [
                "mp4",
                "mov",
                "avi",
                "mkv",
                "webm",
                "mpeg",
                "mpg"
            ];

            const filename =
                file.name || "";

            const extension =
                filename
                    .split(".")
                    .pop()
                    .toLowerCase();

            if (
                !allowedExtensions.includes(
                    extension
                )
            ) {

                throw new Error(
                    "Unsupported video format. " +
                    "Please select MP4, MOV, AVI, MKV, " +
                    "WEBM, MPEG or MPG."
                );
            }

            if (file.size <= 0) {

                throw new Error(
                    "The selected video is empty."
                );
            }

            if (file.size > maxSize) {

                throw new Error(
                    "The selected video is too large."
                );
            }
        }


        /*
        ========================================================
        Video metadata loaded
        ========================================================
        */

        videoPreview.addEventListener(
            "loadedmetadata",
            function () {

                try {

                    videoWidth =
                        videoPreview.videoWidth;

                    videoHeight =
                        videoPreview.videoHeight;

                    if (
                        !videoWidth ||
                        !videoHeight
                    ) {

                        throw new Error(
                            "Unable to determine video dimensions."
                        );
                    }

                    originalDimensions.textContent =
                        `${videoWidth} × ${videoHeight} px`;

                    setupInitialCrop();

                    updateCropBox();

                    updateDisplays();

                    updateHiddenInputs();

                } catch (error) {

                    console.error(
                        "Video metadata error:",
                        error
                    );

                    alert(
                        "Unable to read the video dimensions."
                    );
                }

            }
        );


        /*
        ========================================================
        Initial crop
        ========================================================
        */

        function setupInitialCrop() {

            /*
            Start with approximately 80% of the video.
            */

            let width =
                Math.round(
                    videoWidth * 0.8
                );

            let height =
                Math.round(
                    videoHeight * 0.8
                );

            width = makeEven(width);
            height = makeEven(height);

            crop.width = width;
            crop.height = height;

            crop.x =
                Math.round(
                    (videoWidth - width) / 2
                );

            crop.y =
                Math.round(
                    (videoHeight - height) / 2
                );

            crop.x = Math.max(
                0,
                crop.x
            );

            crop.y = Math.max(
                0,
                crop.y
            );
        }


        /*
        ========================================================
        Ratio buttons
        ========================================================
        */

        ratioButtons.forEach(
            function (button) {

                button.addEventListener(
                    "click",
                    function () {

                        try {

                            ratioButtons.forEach(
                                function (item) {
                                    item.classList.remove(
                                        "active"
                                    );
                                }
                            );

                            button.classList.add(
                                "active"
                            );

                            activeRatio =
                                button.dataset.ratio;

                            cropRatioInput.value =
                                activeRatio;

                            if (
                                activeRatio ===
                                "custom"
                            ) {

                                customSection.style.display =
                                    "block";

                                applyCustomCrop();

                            } else {

                                customSection.style.display =
                                    "none";

                                if (
                                    activeRatio ===
                                    "free"
                                ) {

                                    updateDisplays();
                                    updateHiddenInputs();

                                    return;
                                }

                                applyAspectRatio(
                                    activeRatio
                                );
                            }

                            updateCropBox();

                            updateDisplays();

                            updateHiddenInputs();

                        } catch (error) {

                            console.error(
                                "Ratio selection error:",
                                error
                            );
                        }

                    }
                );

            }
        );


        /*
        ========================================================
        Apply aspect ratio
        ========================================================
        */

        function applyAspectRatio(
            ratioName
        ) {

            const ratio =
                RATIOS[ratioName];

            if (!ratio) {

                throw new Error(
                    "Invalid crop ratio."
                );
            }

            const targetRatio =
                ratio.width /
                ratio.height;

            /*
            Keep the crop area reasonably large.
            */

            let width =
                Math.round(
                    videoWidth * 0.7
                );

            let height =
                Math.round(
                    width / targetRatio
                );

            /*
            If height is too large, calculate
            from video height instead.
            */

            if (
                height > videoHeight * 0.8
            ) {

                height =
                    Math.round(
                        videoHeight * 0.8
                    );

                width =
                    Math.round(
                        height * targetRatio
                    );
            }

            /*
            If width is too large.
            */

            if (
                width > videoWidth
            ) {

                width =
                    videoWidth;

                height =
                    Math.round(
                        width / targetRatio
                    );
            }

            /*
            If height is too large.
            */

            if (
                height > videoHeight
            ) {

                height =
                    videoHeight;

                width =
                    Math.round(
                        height * targetRatio
                    );
            }

            width =
                makeEven(width);

            height =
                makeEven(height);

            if (
                width <= 0 ||
                height <= 0
            ) {

                throw new Error(
                    "Unable to calculate crop dimensions."
                );
            }

            crop.width = width;
            crop.height = height;

            crop.x =
                Math.round(
                    (videoWidth - width) / 2
                );

            crop.y =
                Math.round(
                    (videoHeight - height) / 2
                );

            clampCrop();
        }


        /*
        ========================================================
        Custom crop
        ========================================================
        */

        if (customWidthInput) {

            customWidthInput.addEventListener(
                "input",
                function () {

                    try {

                        if (
                            activeRatio !==
                            "custom"
                        ) {
                            return;
                        }

                        applyCustomCrop();

                        updateCropBox();

                        updateDisplays();

                        updateHiddenInputs();

                    } catch (error) {

                        console.error(
                            "Custom width error:",
                            error
                        );
                    }

                }
            );
        }


        if (customHeightInput) {

            customHeightInput.addEventListener(
                "input",
                function () {

                    try {

                        if (
                            activeRatio !==
                            "custom"
                        ) {
                            return;
                        }

                        applyCustomCrop();

                        updateCropBox();

                        updateDisplays();

                        updateHiddenInputs();

                    } catch (error) {

                        console.error(
                            "Custom height error:",
                            error
                        );
                    }

                }
            );
        }


        function applyCustomCrop() {

            if (
                !customWidthInput ||
                !customHeightInput
            ) {
                return;
            }

            const width =
                parseInt(
                    customWidthInput.value,
                    10
                );

            const height =
                parseInt(
                    customHeightInput.value,
                    10
                );

            if (
                !Number.isFinite(width) ||
                !Number.isFinite(height)
            ) {
                return;
            }

            if (
                width <= 0 ||
                height <= 0
            ) {
                return;
            }

            if (
                width > videoWidth ||
                height > videoHeight
            ) {

                return;
            }

            crop.width =
                makeEven(width);

            crop.height =
                makeEven(height);

            crop.x =
                Math.round(
                    (videoWidth -
                        crop.width) / 2
                );

            crop.y =
                Math.round(
                    (videoHeight -
                        crop.height) / 2
                );

            clampCrop();
        }


        /*
        ========================================================
        Drag crop box
        ========================================================
        */

        cropBox.addEventListener(
            "mousedown",
            startDrag
        );

        cropBox.addEventListener(
            "touchstart",
            startDrag,
            {
                passive: false
            }
        );


        function startDrag(event) {

            try {

                /*
                Don't start dragging when the user
                clicked a resize handle.
                */

                if (
                    event.target.classList.contains(
                        "handle"
                    )
                ) {
                    return;
                }

                event.preventDefault();

                const point =
                    getPointerPosition(
                        event
                    );

                interaction = {
                    type: "drag",
                    startX: point.x,
                    startY: point.y,
                    originalX: crop.x,
                    originalY: crop.y
                };

                document.addEventListener(
                    "mousemove",
                    handlePointerMove
                );

                document.addEventListener(
                    "mouseup",
                    stopInteraction
                );

                document.addEventListener(
                    "touchmove",
                    handlePointerMove,
                    {
                        passive: false
                    }
                );

                document.addEventListener(
                    "touchend",
                    stopInteraction
                );

            } catch (error) {

                console.error(
                    "Crop drag error:",
                    error
                );
            }
        }


        /*
        ========================================================
        Resize handles
        ========================================================
        */

        const handles =
            document.querySelectorAll(
                ".handle"
            );

        handles.forEach(
            function (handle) {

                handle.addEventListener(
                    "mousedown",
                    function (event) {

                        startResize(
                            event,
                            handle.dataset.handle
                        );

                    }
                );

                handle.addEventListener(
                    "touchstart",
                    function (event) {

                        startResize(
                            event,
                            handle.dataset.handle
                        );

                    },
                    {
                        passive: false
                    }
                );

            }
        );


        function startResize(
            event,
            handle
        ) {

            try {

                event.preventDefault();

                event.stopPropagation();

                const point =
                    getPointerPosition(
                        event
                    );

                interaction = {
                    type: "resize",
                    handle: handle,
                    startX: point.x,
                    startY: point.y,

                    originalX: crop.x,
                    originalY: crop.y,

                    originalWidth:
                        crop.width,

                    originalHeight:
                        crop.height
                };

                document.addEventListener(
                    "mousemove",
                    handlePointerMove
                );

                document.addEventListener(
                    "mouseup",
                    stopInteraction
                );

                document.addEventListener(
                    "touchmove",
                    handlePointerMove,
                    {
                        passive: false
                    }
                );

                document.addEventListener(
                    "touchend",
                    stopInteraction
                );

            } catch (error) {

                console.error(
                    "Crop resize error:",
                    error
                );
            }
        }


        /*
        ========================================================
        Pointer move
        ========================================================
        */

        function handlePointerMove(
            event
        ) {

            try {

                if (!interaction) {
                    return;
                }

                event.preventDefault();

                const point =
                    getPointerPosition(
                        event
                    );

                const editorRect =
                    videoEditor.getBoundingClientRect();

                if (
                    editorRect.width <= 0 ||
                    editorRect.height <= 0
                ) {
                    return;
                }

                /*
                Convert screen movement to actual
                video pixels.
                */

                const scaleX =
                    videoWidth /
                    editorRect.width;

                const scaleY =
                    videoHeight /
                    editorRect.height;

                const deltaX =
                    (point.x -
                        interaction.startX) *
                    scaleX;

                const deltaY =
                    (point.y -
                        interaction.startY) *
                    scaleY;


                if (
                    interaction.type ===
                    "drag"
                ) {

                    crop.x =
                        interaction.originalX +
                        deltaX;

                    crop.y =
                        interaction.originalY +
                        deltaY;

                    clampCrop();

                } else if (
                    interaction.type ===
                    "resize"
                ) {

                    resizeCrop(
                        deltaX,
                        deltaY
                    );
                }

                updateCropBox();

                updateDisplays();

                updateHiddenInputs();

            } catch (error) {

                console.error(
                    "Crop interaction error:",
                    error
                );
            }
        }


        /*
        ========================================================
        Resize crop
        ========================================================
        */

        function resizeCrop(
            deltaX,
            deltaY
        ) {

            const original =
                interaction;

            let newX =
                original.originalX;

            let newY =
                original.originalY;

            let newWidth =
                original.originalWidth;

            let newHeight =
                original.originalHeight;


            /*
            ====================================================
            FREE CROP
            ====================================================
            */

            if (
                activeRatio === "free"
            ) {

                switch (
                    original.handle
                ) {

                    case "nw":

                        newX =
                            original.originalX +
                            deltaX;

                        newY =
                            original.originalY +
                            deltaY;

                        newWidth =
                            original.originalWidth -
                            deltaX;

                        newHeight =
                            original.originalHeight -
                            deltaY;

                        break;


                    case "ne":

                        newY =
                            original.originalY +
                            deltaY;

                        newWidth =
                            original.originalWidth +
                            deltaX;

                        newHeight =
                            original.originalHeight -
                            deltaY;

                        break;


                    case "sw":

                        newX =
                            original.originalX +
                            deltaX;

                        newWidth =
                            original.originalWidth -
                            deltaX;

                        newHeight =
                            original.originalHeight +
                            deltaY;

                        break;


                    case "se":

                        newWidth =
                            original.originalWidth +
                            deltaX;

                        newHeight =
                            original.originalHeight +
                            deltaY;

                        break;
                }

            }

            /*
            ====================================================
            RATIO LOCKED CROP
            ====================================================
            */

            else {

                let ratio = null;

                if (
                    activeRatio ===
                    "custom"
                ) {

                    if (
                        original.originalHeight
                        <= 0
                    ) {
                        return;
                    }

                    ratio =
                        original.originalWidth /
                        original.originalHeight;

                } else {

                    const ratioData =
                        RATIOS[
                            activeRatio
                        ];

                    if (!ratioData) {
                        return;
                    }

                    ratio =
                        ratioData.width /
                        ratioData.height;
                }


                switch (
                    original.handle
                ) {

                    case "se":

                        newWidth =
                            original.originalWidth +
                            deltaX;

                        newHeight =
                            newWidth /
                            ratio;

                        break;


                    case "sw":

                        newWidth =
                            original.originalWidth -
                            deltaX;

                        newHeight =
                            newWidth /
                            ratio;

                        newX =
                            original.originalX +
                            (
                                original.originalWidth -
                                newWidth
                            );

                        break;


                    case "ne":

                        newWidth =
                            original.originalWidth +
                            deltaX;

                        newHeight =
                            newWidth /
                            ratio;

                        newY =
                            original.originalY +
                            (
                                original.originalHeight -
                                newHeight
                            );

                        break;


                    case "nw":

                        newWidth =
                            original.originalWidth -
                            deltaX;

                        newHeight =
                            newWidth /
                            ratio;

                        newX =
                            original.originalX +
                            (
                                original.originalWidth -
                                newWidth
                            );

                        newY =
                            original.originalY +
                            (
                                original.originalHeight -
                                newHeight
                            );

                        break;
                }
            }


            /*
            ====================================================
            Minimum size
            ====================================================
            */

            const minimumSize = 20;

            if (
                newWidth <
                minimumSize
            ) {

                newWidth =
                    minimumSize;
            }

            if (
                newHeight <
                minimumSize
            ) {

                newHeight =
                    minimumSize;
            }


            /*
            ====================================================
            Convert to integers
            ====================================================
            */

            newWidth =
                Math.round(
                    newWidth
                );

            newHeight =
                Math.round(
                    newHeight
                );

            newX =
                Math.round(
                    newX
                );

            newY =
                Math.round(
                    newY
                );


            /*
            ====================================================
            Keep crop inside video
            ====================================================
            */

            if (
                newX < 0
            ) {

                newWidth +=
                    newX;

                newX = 0;
            }

            if (
                newY < 0
            ) {

                newHeight +=
                    newY;

                newY = 0;
            }

            if (
                newX +
                newWidth >
                videoWidth
            ) {

                newWidth =
                    videoWidth -
                    newX;
            }

            if (
                newY +
                newHeight >
                videoHeight
            ) {

                newHeight =
                    videoHeight -
                    newY;
            }


            /*
            ====================================================
            Apply
            ====================================================
            */

            crop.x =
                Math.max(
                    0,
                    Math.round(newX)
                );

            crop.y =
                Math.max(
                    0,
                    Math.round(newY)
                );

            crop.width =
                Math.max(
                    minimumSize,
                    Math.round(newWidth)
                );

            crop.height =
                Math.max(
                    minimumSize,
                    Math.round(newHeight)
                );


            /*
            For ratio modes, make sure the final
            crop still follows the ratio.
            */

            if (
                activeRatio !== "free" &&
                activeRatio !== "custom"
            ) {

                enforceRatio();
            }

            clampCrop();
        }


        /*
        ========================================================
        Enforce ratio
        ========================================================
        */

        function enforceRatio() {

            const ratioData =
                RATIOS[
                    activeRatio
                ];

            if (!ratioData) {
                return;
            }

            const targetRatio =
                ratioData.width /
                ratioData.height;

            let width =
                crop.width;

            let height =
                Math.round(
                    width /
                    targetRatio
                );

            if (
                height >
                videoHeight
            ) {

                height =
                    videoHeight;

                width =
                    Math.round(
                        height *
                        targetRatio
                    );
            }

            if (
                width >
                videoWidth
            ) {

                width =
                    videoWidth;

                height =
                    Math.round(
                        width /
                        targetRatio
                    );
            }

            crop.width =
                makeEven(width);

            crop.height =
                makeEven(height);

            clampCrop();
        }


        /*
        ========================================================
        Stop interaction
        ========================================================
        */

        function stopInteraction() {

            interaction = null;

            document.removeEventListener(
                "mousemove",
                handlePointerMove
            );

            document.removeEventListener(
                "mouseup",
                stopInteraction
            );

            document.removeEventListener(
                "touchmove",
                handlePointerMove
            );

            document.removeEventListener(
                "touchend",
                stopInteraction
            );
        }


        /*
        ========================================================
        Get pointer position
        ========================================================
        */

        function getPointerPosition(
            event
        ) {

            let clientX;
            let clientY;

            if (
                event.touches &&
                event.touches.length
            ) {

                clientX =
                    event.touches[0].clientX;

                clientY =
                    event.touches[0].clientY;

            } else {

                clientX =
                    event.clientX;

                clientY =
                    event.clientY;
            }

            const rect =
                videoEditor.getBoundingClientRect();

            return {
                x:
                    clientX -
                    rect.left,

                y:
                    clientY -
                    rect.top
            };
        }


        /*
        ========================================================
        Clamp crop
        ========================================================
        */

        function clampCrop() {

            crop.width =
                Math.min(
                    crop.width,
                    videoWidth
                );

            crop.height =
                Math.min(
                    crop.height,
                    videoHeight
                );

            crop.width =
                Math.max(
                    2,
                    crop.width
                );

            crop.height =
                Math.max(
                    2,
                    crop.height
                );

            crop.x =
                Math.max(
                    0,
                    crop.x
                );

            crop.y =
                Math.max(
                    0,
                    crop.y
                );

            if (
                crop.x +
                crop.width >
                videoWidth
            ) {

                crop.x =
                    videoWidth -
                    crop.width;
            }

            if (
                crop.y +
                crop.height >
                videoHeight
            ) {

                crop.y =
                    videoHeight -
                    crop.height;
            }

            crop.x =
                Math.max(
                    0,
                    crop.x
                );

            crop.y =
                Math.max(
                    0,
                    crop.y
                );
        }


        /*
        ========================================================
        Update crop rectangle on screen
        ========================================================
        */

        function updateCropBox() {

            if (
                !videoWidth ||
                !videoHeight
            ) {
                return;
            }

            const rect =
                videoPreview.getBoundingClientRect();

            const editorRect =
                videoEditor.getBoundingClientRect();

            if (
                !rect.width ||
                !rect.height
            ) {
                return;
            }

            const scaleX =
                rect.width /
                videoWidth;

            const scaleY =
                rect.height /
                videoHeight;

            cropBox.style.left =
                `${crop.x * scaleX}px`;

            cropBox.style.top =
                `${crop.y * scaleY}px`;

            cropBox.style.width =
                `${crop.width * scaleX}px`;

            cropBox.style.height =
                `${crop.height * scaleY}px`;
        }


        /*
        ========================================================
        Update display values
        ========================================================
        */

        function updateDisplays() {

            cropWidthDisplay.textContent =
                `${Math.round(crop.width)} px`;

            cropHeightDisplay.textContent =
                `${Math.round(crop.height)} px`;

            cropXDisplay.textContent =
                `${Math.round(crop.x)} px`;

            cropYDisplay.textContent =
                `${Math.round(crop.y)} px`;

            rotationDisplay.textContent =
                `${rotation}°`;
        }


        /*
        ========================================================
        Update hidden form inputs
        ========================================================
        */

        function updateHiddenInputs() {

            cropRatioInput.value =
                activeRatio;

            cropXInput.value =
                Math.round(crop.x);

            cropYInput.value =
                Math.round(crop.y);

            cropWidthInput.value =
                Math.round(crop.width);

            cropHeightInput.value =
                Math.round(crop.height);

            rotationInput.value =
                rotation;
        }


        /*
        ========================================================
        Rotation
        ========================================================
        */

        rotateLeft.addEventListener(
            "click",
            function () {

                try {

                    rotation -= 90;

                    if (
                        rotation < 0
                    ) {
                        rotation = 270;
                    }

                    applyRotationPreview();

                    updateDisplays();

                    updateHiddenInputs();

                } catch (error) {

                    console.error(
                        "Rotate left error:",
                        error
                    );
                }

            }
        );


        rotateRight.addEventListener(
            "click",
            function () {

                try {

                    rotation += 90;

                    if (
                        rotation >= 360
                    ) {
                        rotation = 0;
                    }

                    applyRotationPreview();

                    updateDisplays();

                    updateHiddenInputs();

                } catch (error) {

                    console.error(
                        "Rotate right error:",
                        error
                    );
                }

            }
        );


        /*
        ========================================================
        Rotation preview
        ========================================================
        */

        function applyRotationPreview() {

            videoPreview.style.transform =
                `rotate(${rotation}deg)`;

            /*
            Keep the editor visually centered.
            */

            if (
                rotation === 90 ||
                rotation === 270
            ) {

                videoPreview.style.maxHeight =
                    "500px";

            } else {

                videoPreview.style.maxHeight =
                    "600px";
            }

            /*
            Rotation is mainly a preview here.
            The backend must apply the actual
            rotation to the generated video.
            */
        }


        /*
        ========================================================
        Window resize
        ========================================================
        */

        window.addEventListener(
            "resize",
            function () {

                try {

                    updateCropBox();

                } catch (error) {

                    console.error(
                        "Editor resize error:",
                        error
                    );
                }

            }
        );


        /*
        ========================================================
        Submit validation
        ========================================================
        */

        form.addEventListener(
            "submit",
            function (event) {

                try {

                    /*
                    Make sure a video exists.
                    */

                    if (
                        !videoInput.files ||
                        !videoInput.files.length
                    ) {

                        event.preventDefault();

                        alert(
                            "Please select a video."
                        );

                        return;
                    }


                    /*
                    Make sure video metadata exists.
                    */

                    if (
                        !videoWidth ||
                        !videoHeight
                    ) {

                        event.preventDefault();

                        alert(
                            "The video is still loading. " +
                            "Please wait and try again."
                        );

                        return;
                    }


                    /*
                    Clamp one final time.
                    */

                    clampCrop();


                    /*
                    Validate dimensions.
                    */

                    if (
                        crop.width <= 0 ||
                        crop.height <= 0
                    ) {

                        event.preventDefault();

                        alert(
                            "Please select a valid crop area."
                        );

                        return;
                    }


                    /*
                    Crop cannot exceed video.
                    */

                    if (
                        crop.x < 0 ||
                        crop.y < 0 ||
                        crop.x + crop.width >
                            videoWidth ||
                        crop.y + crop.height >
                            videoHeight
                    ) {

                        event.preventDefault();

                        alert(
                            "The crop area is outside the video."
                        );

                        return;
                    }


                    /*
                    Custom validation.
                    */

                    if (
                        activeRatio ===
                        "custom"
                    ) {

                        const customWidth =
                            parseInt(
                                customWidthInput.value,
                                10
                            );

                        const customHeight =
                            parseInt(
                                customHeightInput.value,
                                10
                            );

                        if (
                            !Number.isFinite(
                                customWidth
                            ) ||
                            !Number.isFinite(
                                customHeight
                            )
                        ) {

                            event.preventDefault();

                            alert(
                                "Please enter valid custom width and height."
                            );

                            return;
                        }

                        if (
                            customWidth >
                                videoWidth ||
                            customHeight >
                                videoHeight
                        ) {

                            event.preventDefault();

                            alert(
                                "Custom crop dimensions cannot " +
                                "be larger than the original video."
                            );

                            return;
                        }
                    }


                    /*
                    Update hidden values before submit.
                    */

                    updateHiddenInputs();


                    /*
                    Show spinner.
                    */

                    if (
                        processingOverlay
                    ) {

                        processingOverlay.style.display =
                            "flex";
                    }


                    /*
                    Prevent double submission.
                    */

                    if (
                        cropSubmit
                    ) {

                        cropSubmit.disabled =
                            true;

                        cropSubmit.textContent =
                            "Processing...";
                    }

                } catch (error) {

                    event.preventDefault();

                    console.error(
                        "Crop form submission error:",
                        error
                    );

                    if (
                        processingOverlay
                    ) {

                        processingOverlay.style.display =
                            "none";
                    }

                    if (
                        cropSubmit
                    ) {

                        cropSubmit.disabled =
                            false;

                        cropSubmit.textContent =
                            "Crop Video";
                    }

                    alert(
                        "Unable to submit the crop request."
                    );
                }

            }
        );


        /*
        ========================================================
        Reset editor
        ========================================================
        */

        function resetEditor() {

            try {

                if (objectUrl) {

                    URL.revokeObjectURL(
                        objectUrl
                    );

                    objectUrl = null;
                }

                videoPreview.removeAttribute(
                    "src"
                );

                videoPreview.load();

                editorSection.style.display =
                    "none";

                videoWidth = 0;
                videoHeight = 0;

                crop = {
                    x: 0,
                    y: 0,
                    width: 0,
                    height: 0
                };

                rotation = 0;

                activeRatio =
                    "free";

                cropRatioInput.value =
                    "free";

                cropXInput.value =
                    "0";

                cropYInput.value =
                    "0";

                cropWidthInput.value =
                    "0";

                cropHeightInput.value =
                    "0";

                rotationInput.value =
                    "0";

                videoPreview.style.transform =
                    "rotate(0deg)";

            } catch (error) {

                console.error(
                    "Crop editor reset error:",
                    error
                );
            }
        }


        /*
        ========================================================
        Make number even
        ========================================================
        */

        function makeEven(value) {

            value =
                Math.round(
                    Number(value)
                );

            if (
                value % 2 !== 0
            ) {
                value -= 1;
            }

            return Math.max(
                2,
                value
            );
        }


        /*
        ========================================================
        Cleanup object URL
        ========================================================
        */

        window.addEventListener(
            "beforeunload",
            function () {

                try {

                    if (objectUrl) {

                        URL.revokeObjectURL(
                            objectUrl
                        );
                    }

                } catch (error) {

                    console.error(
                        "Object URL cleanup failed:",
                        error
                    );
                }

            }
        );


        /*
        ========================================================
        Initial state
        ========================================================
        */

        editorSection.style.display =
            "none";

    }

})();