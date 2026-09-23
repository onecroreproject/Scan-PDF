document.addEventListener("DOMContentLoaded", function () {

    "use strict";

    /* =========================================================
       ELEMENTS
    ========================================================= */

    const uploadScreen = document.getElementById("uploadScreen");
    const editorScreen = document.getElementById("editorScreen");
    const exportingScreen = document.getElementById("exportingScreen");
    const resultScreen = document.getElementById("resultScreen");

    const uploadButton = document.getElementById("uploadButton");

    const rotateForm = document.getElementById("rotateForm");

    const videoInput = document.getElementById("id_video");

    const videoPreview = document.getElementById("videoPreview");
    const resultVideo = document.getElementById("resultVideo");

    const rotationField = document.getElementById("id_rotation");
    const outputFormatField = document.getElementById("id_output_format");

    const rotationBadge = document.getElementById("rotationBadge");

    const rotateLeftButton =
        document.getElementById("rotateLeftButton");

    const rotateRightButton =
        document.getElementById("rotateRightButton");

    const rotate180Button =
        document.getElementById("rotate180Button");

    const resetRotationButton =
        document.getElementById("resetRotationButton");

    const rotationButtons =
        document.querySelectorAll(".rotate-button");

    const outputFormatSelect =
        document.getElementById("outputFormatSelect");

    const exportButton =
        document.getElementById("exportButton");

    const downloadButton =
        document.getElementById("downloadButton");

    const editAgainButton =
        document.getElementById("editAgainButton");


    /* =========================================================
       STATE
    ========================================================= */

    let selectedRotation = 0;
    let previewObjectUrl = null;

    const ALLOWED_ROTATIONS = [
        0,
        90,
        180,
        270
    ];

    const ALLOWED_FORMATS = [
        "mp4",
        "mov",
        "webm",
        "avi",
        "mkv",
        "m4v",
        "gif",
        "mpeg",
        "flv",
        "wmv"
    ];


    /* =========================================================
       SCREEN MANAGEMENT
    ========================================================= */

    function showScreen(screen) {

        uploadScreen.classList.add("hidden");
        editorScreen.classList.add("hidden");
        exportingScreen.classList.add("hidden");
        resultScreen.classList.add("hidden");

        screen.classList.remove("hidden");

        window.scrollTo({
            top: 0,
            behavior: "smooth"
        });
    }


    /* =========================================================
       MESSAGES
    ========================================================= */

    function showMessage(message, type = "error") {

        const oldMessages =
            document.querySelectorAll(".rotate-message");

        oldMessages.forEach(function (item) {
            item.remove();
        });

        const messageElement =
            document.createElement("div");

        messageElement.className =
            "rotate-message alert alert-" + type;

        messageElement.textContent = message;

        document
            .querySelector(".rotate-page")
            .prepend(messageElement);

        setTimeout(function () {

            if (messageElement) {
                messageElement.style.opacity = "0";
                messageElement.style.transition =
                    "opacity .3s ease";

                setTimeout(function () {
                    messageElement.remove();
                }, 300);
            }

        }, 5000);
    }


    /* =========================================================
       FILE VALIDATION
    ========================================================= */

    function validateVideoFile(file) {

        if (!file) {
            showMessage("Please select a video file.");
            return false;
        }

        if (!file.type.startsWith("video/")) {
            showMessage("Please select a valid video file.");
            return false;
        }

        if (file.size <= 0) {
            showMessage("The selected video file is empty.");
            return false;
        }

        return true;
    }


    /* =========================================================
       UPLOAD BUTTON
    ========================================================= */

    uploadButton.addEventListener("click", function () {

        if (!videoInput) {
            showMessage("Video upload field was not found.");
            return;
        }

        videoInput.click();
    });


    /* =========================================================
       FILE SELECTED
    ========================================================= */

    if (videoInput) {

        videoInput.addEventListener("change", function () {

            const file = this.files[0];

            if (!validateVideoFile(file)) {
                return;
            }

            loadSelectedVideo(file);

        });

    }


    /* =========================================================
       LOAD VIDEO
    ========================================================= */

    function loadSelectedVideo(file) {

        if (previewObjectUrl) {

            URL.revokeObjectURL(
                previewObjectUrl
            );

            previewObjectUrl = null;
        }

        previewObjectUrl =
            URL.createObjectURL(file);

        videoPreview.src =
            previewObjectUrl;

        videoPreview.load();

        selectedRotation = 0;

        updateRotation();

        exportButton.disabled = true;

        showScreen(editorScreen);
    }


    /* =========================================================
       VIDEO METADATA
    ========================================================= */

    videoPreview.addEventListener(
        "loadedmetadata",
        function () {

            exportButton.disabled = false;

        }
    );


    /* =========================================================
       ROTATION
    ========================================================= */

    function normalizeRotation(rotation) {

        rotation =
            parseInt(rotation, 10);

        if (isNaN(rotation)) {
            rotation = 0;
        }

        rotation =
            rotation % 360;

        if (rotation < 0) {
            rotation += 360;
        }

        return rotation;
    }


    function setRotation(rotation) {

        rotation =
            normalizeRotation(rotation);

        if (!ALLOWED_ROTATIONS.includes(rotation)) {
            rotation = 0;
        }

        selectedRotation = rotation;

        updateRotation();
    }


    function updateRotation() {

        /*
         * Rotate the preview visually.
         */
        videoPreview.style.transform =
            `rotate(${selectedRotation}deg)`;


        /*
         * Update Django hidden form field.
         */
        if (rotationField) {
            rotationField.value =
                selectedRotation;
        }


        /*
         * Update badge.
         */
        rotationBadge.textContent =
            `${selectedRotation}°`;


        /*
         * Active button.
         */
        rotationButtons.forEach(function (button) {

            const buttonRotation =
                parseInt(
                    button.dataset.rotation,
                    10
                );

            button.classList.toggle(
                "active",
                buttonRotation === selectedRotation
            );

        });
    }


    /* =========================================================
       ROTATE LEFT
       ========================================================= */

    rotateLeftButton.addEventListener(
        "click",
        function () {

            setRotation(
                selectedRotation === 0
                    ? 270
                    : selectedRotation - 90
            );

        }
    );


    /* =========================================================
       ROTATE RIGHT
       ========================================================= */

    rotateRightButton.addEventListener(
        "click",
        function () {

            setRotation(
                (selectedRotation + 90) % 360
            );

        }
    );


    /* =========================================================
       ROTATE 180
       ========================================================= */

    rotate180Button.addEventListener(
        "click",
        function () {

            setRotation(
                selectedRotation === 180
                    ? 0
                    : 180
            );

        }
    );


    /* =========================================================
       RESET
    ========================================================= */

    resetRotationButton.addEventListener(
        "click",
        function () {

            setRotation(0);

        }
    );


    /* =========================================================
       GENERIC ROTATION BUTTON SUPPORT
       ========================================================= */

    rotationButtons.forEach(function (button) {

        button.addEventListener(
            "click",
            function () {

                const rotation =
                    parseInt(
                        button.dataset.rotation,
                        10
                    );

                if (
                    !isNaN(rotation) &&
                    ALLOWED_ROTATIONS.includes(rotation)
                ) {
                    setRotation(rotation);
                }

            }
        );

    });


    /* =========================================================
       OUTPUT FORMAT
    ========================================================= */

    function updateOutputFormat() {

        if (!outputFormatSelect) {
            return;
        }

        let format =
            outputFormatSelect.value
                .toLowerCase()
                .trim();

        if (!ALLOWED_FORMATS.includes(format)) {

            format = "mp4";

            outputFormatSelect.value =
                format;
        }

        if (outputFormatField) {

            outputFormatField.value =
                format;

        }

    }


    outputFormatSelect.addEventListener(
        "change",
        updateOutputFormat
    );


    updateOutputFormat();


    /* =========================================================
       FORM SUBMIT
    ========================================================= */

    rotateForm.addEventListener(
        "submit",
        async function (event) {

            event.preventDefault();


            /*
             * Make sure a video exists.
             */
            if (
                !videoInput ||
                !videoInput.files ||
                !videoInput.files[0]
            ) {

                showMessage(
                    "Please select a video first."
                );

                return;
            }


            /*
             * Validate video.
             */
            const file =
                videoInput.files[0];

            if (!validateVideoFile(file)) {
                return;
            }


            /*
             * Make sure rotation is valid.
             */
            if (
                !ALLOWED_ROTATIONS.includes(
                    selectedRotation
                )
            ) {

                showMessage(
                    "Invalid rotation selected."
                );

                return;
            }


            /*
             * Update hidden fields.
             */
            if (rotationField) {

                rotationField.value =
                    selectedRotation;

            }


            updateOutputFormat();


            /*
             * Build FormData.
             */
            const formData =
                new FormData(rotateForm);


            /*
             * Show processing screen.
             */
            showScreen(exportingScreen);


            exportButton.disabled = true;


            try {

                const response =
                    await fetch(
                        rotateForm.action,
                        {
                            method: "POST",

                            body: formData,

                            headers: {
                                "X-Requested-With":
                                    "XMLHttpRequest"
                            },

                            credentials: "same-origin"
                        }
                    );


                /*
                 * Try to parse JSON.
                 */
                let data;

                try {

                    data =
                        await response.json();

                } catch (jsonError) {

                    throw new Error(
                        "The server returned an invalid response."
                    );

                }


                /*
                 * Backend returned an error.
                 */
                if (
                    !response.ok ||
                    !data.success
                ) {

                    throw new Error(
                        data.error ||
                        data.message ||
                        "Video processing failed."
                    );

                }


                /*
                 * IMPORTANT:
                 * Use the CURRENT response URL.
                 *
                 * Do not display an old download URL.
                 */
                const videoUrl =
                    data.video_url ||
                    data.url;

                const downloadUrl =
                    data.download_url ||
                    data.video_url ||
                    data.url;


                if (!videoUrl) {

                    throw new Error(
                        "Processing completed but no video URL was returned."
                    );

                }


                /*
                 * Set result video.
                 */
                resultVideo.src =
                    videoUrl;

                resultVideo.load();


                /*
                 * Set CURRENT download URL.
                 */
                if (downloadUrl) {

                    downloadButton.href =
                        downloadUrl;

                    if (data.filename) {

                        downloadButton.download =
                            data.filename;

                    }

                }


                /*
                 * Only NOW show result.
                 */
                showScreen(resultScreen);


            } catch (error) {

                console.error(
                    "Rotate processing error:",
                    error
                );


                /*
                 * Return to editor if processing failed.
                 */
                showScreen(editorScreen);


                exportButton.disabled = false;


                showMessage(
                    error.message ||
                    "Something went wrong while processing the video."
                );

            }

        }
    );


    /* =========================================================
       EDIT AGAIN
    ========================================================= */

    editAgainButton.addEventListener(
        "click",
        function () {

            /*
             * Stop result video.
             */
            resultVideo.pause();

            resultVideo.removeAttribute("src");

            resultVideo.load();


            /*
             * Reset rotation.
             */
            selectedRotation = 0;

            updateRotation();


            /*
             * Re-enable export.
             */
            exportButton.disabled = false;


            /*
             * Go back to editor.
             */
            showScreen(editorScreen);

        }
    );


    /* =========================================================
       CLEAN OBJECT URL
    ========================================================= */

    window.addEventListener(
        "beforeunload",
        function () {

            if (previewObjectUrl) {

                URL.revokeObjectURL(
                    previewObjectUrl
                );

                previewObjectUrl = null;
            }

        }
    );


    /* =========================================================
       PAUSE VIDEO WHEN SCREEN CHANGES
    ========================================================= */

    document.addEventListener(
        "visibilitychange",
        function () {

            if (document.hidden) {

                videoPreview.pause();
                resultVideo.pause();

            }

        }
    );


    /* =========================================================
       INITIAL STATE
    ========================================================= */

    setRotation(0);

    updateOutputFormat();

    showScreen(uploadScreen);

});