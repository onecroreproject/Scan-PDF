document.addEventListener(
    "DOMContentLoaded",
    function () {

        "use strict";

        try {

            const form =
                document.getElementById(
                    "cropForm"
                );

            const fileInput =
                document.getElementById(
                    "id_video"
                );

            const editor =
                document.getElementById(
                    "videoEditor"
                );

            const video =
                document.getElementById(
                    "videoPreview"
                );

            const cropBox =
                document.getElementById(
                    "cropBox"
                );

            const cropDimensions =
                document.getElementById(
                    "cropDimensions"
                );

            const cropPosition =
                document.getElementById(
                    "cropPosition"
                );

            const originalDimensions =
                document.getElementById(
                    "originalDimensions"
                );

            const editorWrapper =
                document.getElementById(
                    "editorWrapper"
                );

            const liveCanvas =
                document.getElementById(
                    "liveCropPreview"
                );

            const processButton =
                document.getElementById(
                    "processButton"
                );

            const loading =
                document.getElementById(
                    "loading"
                );

            const xInput =
                document.getElementById(
                    "id_x"
                );

            const yInput =
                document.getElementById(
                    "id_y"
                );

            const widthInput =
                document.getElementById(
                    "id_width"
                );

            const heightInput =
                document.getElementById(
                    "id_height"
                );


            /*
             * State
             */

            let videoWidth = 0;
            let videoHeight = 0;

            let crop = {
                x: 0,
                y: 0,
                width: 0,
                height: 0
            };

            let dragState = null;

            let objectUrl = null;


            /*
             * File selected.
             */

            fileInput.addEventListener(
                "change",
                function () {

                    try {

                        const file =
                            fileInput.files[0];

                        if (!file) {
                            return;
                        }


                        if (
                            !file.type.startsWith(
                                "video/"
                            )
                        ) {

                            showValidationError(
                                "Please select a valid video."
                            );

                            fileInput.value = "";

                            return;
                        }


                        if (objectUrl) {

                            URL.revokeObjectURL(
                                objectUrl
                            );
                        }


                        objectUrl =
                            URL.createObjectURL(
                                file
                            );

                        video.src =
                            objectUrl;

                        video.load();


                        video.addEventListener(
                            "loadedmetadata",
                            initializeCrop,
                            {
                                once: true
                            }
                        );

                    } catch (error) {

                        console.error(
                            "Video loading error:",
                            error
                        );

                        showValidationError(
                            "Unable to load video."
                        );
                    }
                }
            );


            /*
             * Initialize crop after metadata.
             */

            function initializeCrop() {

                try {

                    videoWidth =
                        video.videoWidth;

                    videoHeight =
                        video.videoHeight;


                    if (
                        !videoWidth ||
                        !videoHeight
                    ) {

                        showValidationError(
                            "Unable to detect video dimensions."
                        );

                        return;
                    }


                    originalDimensions.textContent =
                        `${videoWidth} × ${videoHeight}`;


                    /*
                     * Initial crop = 80%
                     * of video.
                     */

                    crop.width =
                        Math.floor(
                            videoWidth * 0.8
                        );

                    crop.height =
                        Math.floor(
                            videoHeight * 0.8
                        );

                    crop.x =
                        Math.floor(
                            (videoWidth -
                                crop.width) / 2
                        );

                    crop.y =
                        Math.floor(
                            (videoHeight -
                                crop.height) / 2
                        );


                    editorWrapper.style.display =
                        "block";


                    /*
                     * Wait for video
                     * display dimensions.
                     */

                    requestAnimationFrame(
                        function () {

                            updateCropUI();

                        }
                    );

                } catch (error) {

                    console.error(
                        "Crop initialization error:",
                        error
                    );

                    showValidationError(
                        "Unable to initialize crop editor."
                    );
                }
            }


            /*
             * Get actual displayed
             * video rectangle.
             */

            function getVideoDisplayRect() {

                const rect =
                    video.getBoundingClientRect();

                return {
                    left: rect.left,
                    top: rect.top,
                    width: rect.width,
                    height: rect.height
                };
            }


            /*
             * Convert video pixels
             * to screen pixels.
             */

            function videoToScreenX(
                value
            ) {

                const rect =
                    getVideoDisplayRect();

                return (
                    value /
                    videoWidth *
                    rect.width
                );
            }


            function videoToScreenY(
                value
            ) {

                const rect =
                    getVideoDisplayRect();

                return (
                    value /
                    videoHeight *
                    rect.height
                );
            }


            /*
             * Convert screen pixels
             * to video pixels.
             */

            function screenToVideoX(
                value
            ) {

                const rect =
                    getVideoDisplayRect();

                return (
                    value /
                    rect.width *
                    videoWidth
                );
            }


            function screenToVideoY(
                value
            ) {

                const rect =
                    getVideoDisplayRect();

                return (
                    value /
                    rect.height *
                    videoHeight
                );
            }


            /*
             * Update crop box.
             */

            function updateCropUI() {

                try {

                    const rect =
                        getVideoDisplayRect();


                    const left =
                        videoToScreenX(
                            crop.x
                        );

                    const top =
                        videoToScreenY(
                            crop.y
                        );

                    const width =
                        videoToScreenX(
                            crop.width
                        );

                    const height =
                        videoToScreenY(
                            crop.height
                        );


                    /*
                     * Video element can have
                     * different parent position.
                     */

                    const editorRect =
                        editor.getBoundingClientRect();


                    cropBox.style.left =
                        `${rect.left - editorRect.left + left}px`;

                    cropBox.style.top =
                        `${rect.top - editorRect.top + top}px`;

                    cropBox.style.width =
                        `${width}px`;

                    cropBox.style.height =
                        `${height}px`;


                    /*
                     * Update text.
                     */

                    cropDimensions.textContent =
                        `${Math.round(crop.width)} × ${Math.round(crop.height)}`;


                    cropPosition.textContent =
                        `${Math.round(crop.x)}, ${Math.round(crop.y)}`;


                    /*
                     * Update Django
                     * hidden fields.
                     */

                    xInput.value =
                        Math.round(crop.x);

                    yInput.value =
                        Math.round(crop.y);

                    widthInput.value =
                        Math.round(crop.width);

                    heightInput.value =
                        Math.round(crop.height);


                    updateLivePreview();

                } catch (error) {

                    console.error(
                        "Crop UI update error:",
                        error
                    );
                }
            }


            /*
             * Start drag.
             */

            cropBox.addEventListener(
                "pointerdown",
                function (event) {

                    try {

                        /*
                         * Do not start
                         * move when handle
                         * was clicked.
                         */

                        if (
                            event.target.classList.contains(
                                "handle"
                            )
                        ) {
                            return;
                        }


                        event.preventDefault();

                        cropBox.setPointerCapture(
                            event.pointerId
                        );


                        dragState = {

                            type: "move",

                            startX:
                                event.clientX,

                            startY:
                                event.clientY,

                            originalX:
                                crop.x,

                            originalY:
                                crop.y

                        };

                    } catch (error) {

                        console.error(
                            "Crop drag error:",
                            error
                        );
                    }
                }
            );


            /*
             * Resize handles.
             */

            document
                .querySelectorAll(
                    ".handle"
                )
                .forEach(
                    function (handle) {

                        handle.addEventListener(
                            "pointerdown",
                            function (event) {

                                try {

                                    event.preventDefault();

                                    event.stopPropagation();


                                    const type =
                                        handle.dataset.handle;


                                    handle.setPointerCapture(
                                        event.pointerId
                                    );


                                    dragState = {

                                        type: type,

                                        startX:
                                            event.clientX,

                                        startY:
                                            event.clientY,

                                        originalX:
                                            crop.x,

                                        originalY:
                                            crop.y,

                                        originalWidth:
                                            crop.width,

                                        originalHeight:
                                            crop.height

                                    };

                                } catch (error) {

                                    console.error(
                                        "Resize start error:",
                                        error
                                    );
                                }

                            }
                        );

                    }
                );


            /*
             * Pointer movement.
             */

            document.addEventListener(
                "pointermove",
                function (event) {

                    if (!dragState) {
                        return;
                    }


                    try {

                        const dx =
                            screenToVideoX(
                                event.clientX -
                                dragState.startX
                            );

                        const dy =
                            screenToVideoY(
                                event.clientY -
                                dragState.startY
                            );


                        if (
                            dragState.type ===
                            "move"
                        ) {

                            crop.x =
                                dragState.originalX +
                                dx;

                            crop.y =
                                dragState.originalY +
                                dy;


                            /*
                             * Keep inside video.
                             */

                            crop.x =
                                Math.max(
                                    0,
                                    Math.min(
                                        crop.x,
                                        videoWidth -
                                        crop.width
                                    )
                                );

                            crop.y =
                                Math.max(
                                    0,
                                    Math.min(
                                        crop.y,
                                        videoHeight -
                                        crop.height
                                    )
                                );

                        } else {

                            resizeCrop(
                                dragState.type,
                                dx,
                                dy
                            );
                        }


                        updateCropUI();

                    } catch (error) {

                        console.error(
                            "Pointer move error:",
                            error
                        );
                    }
                }
            );


            /*
             * Stop drag.
             */

            document.addEventListener(
                "pointerup",
                function () {

                    dragState = null;

                }
            );


            /*
             * Resize crop.
             */

            function resizeCrop(
                type,
                dx,
                dy
            ) {

                const minSize = 40;


                if (type === "se") {

                    crop.width =
                        dragState.originalWidth +
                        dx;

                    crop.height =
                        dragState.originalHeight +
                        dy;
                }


                if (type === "sw") {

                    crop.x =
                        dragState.originalX +
                        dx;

                    crop.width =
                        dragState.originalWidth -
                        dx;

                    crop.height =
                        dragState.originalHeight +
                        dy;
                }


                if (type === "ne") {

                    crop.y =
                        dragState.originalY +
                        dy;

                    crop.width =
                        dragState.originalWidth +
                        dx;

                    crop.height =
                        dragState.originalHeight -
                        dy;
                }


                if (type === "nw") {

                    crop.x =
                        dragState.originalX +
                        dx;

                    crop.y =
                        dragState.originalY +
                        dy;

                    crop.width =
                        dragState.originalWidth -
                        dx;

                    crop.height =
                        dragState.originalHeight -
                        dy;
                }


                /*
                 * Minimum size.
                 */

                if (
                    crop.width <
                    minSize
                ) {

                    crop.width =
                        minSize;
                }

                if (
                    crop.height <
                    minSize
                ) {

                    crop.height =
                        minSize;
                }


                /*
                 * Keep inside video.
                 */

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

                    crop.width =
                        videoWidth -
                        crop.x;
                }


                if (
                    crop.y +
                    crop.height >
                    videoHeight
                ) {

                    crop.height =
                        videoHeight -
                        crop.y;
                }
            }


            /*
             * Live crop preview.
             */

            function updateLivePreview() {

                try {

                    if (
                        !video.videoWidth ||
                        !video.videoHeight
                    ) {
                        return;
                    }


                    const context =
                        liveCanvas.getContext(
                            "2d"
                        );


                    /*
                     * Use crop dimensions.
                     */

                    liveCanvas.width =
                        Math.max(
                            1,
                            Math.round(
                                crop.width
                            )
                        );

                    liveCanvas.height =
                        Math.max(
                            1,
                            Math.round(
                                crop.height
                            )
                        );


                    context.clearRect(
                        0,
                        0,
                        liveCanvas.width,
                        liveCanvas.height
                    );


                    context.drawImage(
                        video,

                        crop.x,
                        crop.y,
                        crop.width,
                        crop.height,

                        0,
                        0,
                        liveCanvas.width,
                        liveCanvas.height
                    );

                } catch (error) {

                    console.error(
                        "Live preview error:",
                        error
                    );
                }
            }


            /*
             * Update preview while
             * video plays.
             */

            video.addEventListener(
                "timeupdate",
                function () {

                    updateLivePreview();

                }
            );


            video.addEventListener(
                "seeked",
                function () {

                    updateLivePreview();

                }
            );


            /*
             * Window resize.
             */

            window.addEventListener(
                "resize",
                function () {

                    try {

                        if (
                            videoWidth &&
                            videoHeight
                        ) {

                            updateCropUI();

                        }

                    } catch (error) {

                        console.error(
                            "Resize error:",
                            error
                        );
                    }
                }
            );


            /*
             * Form submit.
             */

            form.addEventListener(
                "submit",
                function (event) {

                    try {

                        if (
                            !videoWidth ||
                            !videoHeight
                        ) {

                            event.preventDefault();

                            showValidationError(
                                "Please select a video first."
                            );

                            return;
                        }


                        /*
                         * Final validation.
                         */

                        if (
                            crop.width <= 0 ||
                            crop.height <= 0
                        ) {

                            event.preventDefault();

                            showValidationError(
                                "Invalid crop area."
                            );

                            return;
                        }


                        if (
                            crop.x < 0 ||
                            crop.y < 0
                        ) {

                            event.preventDefault();

                            showValidationError(
                                "Invalid crop position."
                            );

                            return;
                        }


                        if (
                            crop.x +
                            crop.width >
                            videoWidth
                        ) {

                            event.preventDefault();

                            showValidationError(
                                "Crop area exceeds video width."
                            );

                            return;
                        }


                        if (
                            crop.y +
                            crop.height >
                            videoHeight
                        ) {

                            event.preventDefault();

                            showValidationError(
                                "Crop area exceeds video height."
                            );

                            return;
                        }


                        /*
                         * Make sure hidden
                         * fields contain
                         * latest values.
                         */

                        xInput.value =
                            Math.round(crop.x);

                        yInput.value =
                            Math.round(crop.y);

                        widthInput.value =
                            Math.round(crop.width);

                        heightInput.value =
                            Math.round(crop.height);


                        processButton.disabled =
                            true;

                        processButton.textContent =
                            "Processing...";


                        loading.style.display =
                            "block";


                    } catch (error) {

                        event.preventDefault();

                        console.error(
                            "Form submit error:",
                            error
                        );

                        showValidationError(
                            "Unable to submit crop request."
                        );
                    }
                }
            );

        } catch (error) {

            console.error(
                "Crop editor initialization error:",
                error
            );

        }

    }
);