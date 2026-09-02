document.addEventListener("DOMContentLoaded", function () {


"use strict";

/* =========================================================
   ELEMENTS
   ========================================================= */

const form = document.getElementById("cropForm");
const fileInput = document.getElementById("id_video");
const uploadSection = document.getElementById("uploadSection");
const editorWrapper = document.getElementById("editorWrapper");
const videoEditor = document.getElementById("videoEditor");
const videoPreview = document.getElementById("videoPreview");
const cropBox = document.getElementById("cropBox");
const processButton = document.getElementById("processButton");
const loading = document.getElementById("loading");
const validationPopup = document.getElementById("validationPopup");

const originalDimensions =
    document.getElementById("originalDimensions");

const cropDimensions =
    document.getElementById("cropDimensions");

const cropPosition =
    document.getElementById("cropPosition");

const liveCropPreview =
    document.getElementById("liveCropPreview");

const resultContainer =
    document.getElementById("resultContainer");

const xInput = document.getElementById("id_x");
const yInput = document.getElementById("id_y");
const widthInput = document.getElementById("id_width");
const heightInput = document.getElementById("id_height");


/* =========================================================
   STATE
   ========================================================= */

let videoMetadata = null;
let objectURL = null;
let isProcessing = false;
let dragState = null;

let cropState = {
    x: 0,
    y: 0,
    width: 0,
    height: 0
};


/* =========================================================
   SAFETY CHECK
   ========================================================= */

if (!form || !fileInput) {

    console.error("Crop video form not found.");

    return;
}


/* =========================================================
   INITIAL STATE
   ========================================================= */

if (editorWrapper) {
    editorWrapper.style.display = "none";
}

if (loading) {
    loading.style.display = "none";
}


/* =========================================================
   HIDE OLD RESULT
   ========================================================= */

function hideOldResult() {

    if (resultContainer) {
        resultContainer.style.display = "none";
    }
}


/* =========================================================
   ERROR / WARNING POPUP
   AUTO HIDE AFTER 5 SECONDS
   ========================================================= */

let popupTimer = null;
let popupHideTimer = null;

function showError(message) {

    if (!validationPopup) {
        return;
    }

    /* Cancel previous timers */
    clearTimeout(popupTimer);
    clearTimeout(popupHideTimer);

    /* Set message */
    validationPopup.textContent = message;

    /* Make it visible immediately */
    validationPopup.style.display = "block";

    /* Force animation restart */
    validationPopup.classList.remove("show");

    requestAnimationFrame(function () {

        validationPopup.classList.add("show");

    });


    /* -----------------------------------------------------
       Automatically hide after 5 seconds
       ----------------------------------------------------- */

    popupTimer = setTimeout(function () {

        validationPopup.classList.remove("show");

        popupHideTimer = setTimeout(function () {

            validationPopup.style.display = "none";

        }, 300);

    }, 5000);
}


/* =========================================================
   AUTO HIDE DJANGO MESSAGES
   SUCCESS / ERROR / WARNING / INFO
   AFTER 5 SECONDS
   ========================================================= */

const djangoMessages =
    document.querySelectorAll(
        "#djangoMessages .message, " +
        "#djangoMessages .messages, " +
        ".form-errors"
    );


djangoMessages.forEach(function (element) {

    setTimeout(function () {

        element.style.transition =
            "opacity .4s ease, transform .4s ease";

        element.style.opacity = "0";
        element.style.transform = "translateY(-8px)";

        setTimeout(function () {

            element.remove();

        }, 400);

    }, 5000);

});


/* =========================================================
   FILE VALIDATION
   ========================================================= */

function isValidVideo(file) {

    if (!file) {
        return false;
    }

    if (
        file.type &&
        file.type.startsWith("video/")
    ) {
        return true;
    }

    const extension =
        file.name
            .split(".")
            .pop()
            .toLowerCase();

    const allowed = [
        "mp4",
        "mov",
        "avi",
        "mkv",
        "webm",
        "mpeg",
        "mpg",
        "m4v",
        "3gp"
    ];

    return allowed.includes(extension);
}


/* =========================================================
   VIDEO UPLOAD
   ========================================================= */

fileInput.addEventListener(
    "change",
    function () {

        try {

            const file = fileInput.files[0];

            if (!file) {
                return;
            }


            /* ---------------------------------------------
               Validate video
               --------------------------------------------- */

            if (!isValidVideo(file)) {

                showError(
                    "Please select a valid video file."
                );

                fileInput.value = "";

                return;
            }


            /* ---------------------------------------------
               Hide old result
               --------------------------------------------- */

            hideOldResult();


            /* ---------------------------------------------
               Remove old object URL
               --------------------------------------------- */

            if (objectURL) {

                URL.revokeObjectURL(objectURL);

                objectURL = null;
            }


            /* ---------------------------------------------
               Create browser preview URL
               --------------------------------------------- */

            objectURL =
                URL.createObjectURL(file);

            videoPreview.src = objectURL;

            videoPreview.load();


            /* ---------------------------------------------
               Hide upload section
               --------------------------------------------- */

            if (uploadSection) {

                uploadSection.style.display = "none";
            }


            /* ---------------------------------------------
               Show editor
               --------------------------------------------- */

            if (editorWrapper) {

                editorWrapper.style.display = "block";
            }


            /* ---------------------------------------------
               Reset processing
               --------------------------------------------- */

            isProcessing = false;


            if (processButton) {

                processButton.disabled = false;

                processButton.textContent =
                    "Crop Video";

                processButton.style.display =
                    "block";
            }


            if (loading) {

                loading.style.display = "none";
            }


            /* ---------------------------------------------
               Metadata
               --------------------------------------------- */

            videoPreview.onloadedmetadata =
                function () {

                    const width =
                        videoPreview.videoWidth;

                    const height =
                        videoPreview.videoHeight;


                    if (!width || !height) {

                        showError(
                            "Unable to read video dimensions."
                        );

                        return;
                    }


                    videoMetadata = {
                        width: width,
                        height: height
                    };


                    /* -------------------------------------
                       Display original dimensions
                       ------------------------------------- */

                    if (originalDimensions) {

                        originalDimensions.textContent =
                            `${width} × ${height}`;
                    }


                    /* -------------------------------------
                       Start with full video crop
                       ------------------------------------- */

                    cropState = {

                        x: 0,

                        y: 0,

                        width: width,

                        height: height
                    };


                    updateHiddenInputs();

                    updateCropBox();

                    updateInformation();

                    updateLivePreview();

                };


            /* ---------------------------------------------
               Video load error
               --------------------------------------------- */

            videoPreview.onerror =
                function () {

                    showError(
                        "Unable to load this video. Please select another video."
                    );

                };


        } catch (error) {

            console.error(
                "Video upload error:",
                error
            );

            showError(
                "Something went wrong while loading the video."
            );
        }

    }
);


/* =========================================================
   GET EDITOR SIZE
   ========================================================= */

function getEditorSize() {

    if (!videoEditor) {
        return null;
    }

    const rect =
        videoEditor.getBoundingClientRect();

    if (
        rect.width <= 0 ||
        rect.height <= 0
    ) {
        return null;
    }

    return {
        width: rect.width,
        height: rect.height
    };
}


/* =========================================================
   UPDATE CROP BOX
   ========================================================= */

function updateCropBox() {

    if (
        !cropBox ||
        !videoMetadata
    ) {
        return;
    }


    const editorSize =
        getEditorSize();

    if (!editorSize) {
        return;
    }


    const videoWidth =
        videoPreview.videoWidth;

    const videoHeight =
        videoPreview.videoHeight;

    const containerWidth =
        editorSize.width;

    const containerHeight =
        editorSize.height;

    const videoRatio =
        videoWidth / videoHeight;

    const containerRatio =
        containerWidth / containerHeight;


    let displayedWidth;
    let displayedHeight;
    let offsetX;
    let offsetY;


    if (videoRatio > containerRatio) {

        displayedWidth =
            containerWidth;

        displayedHeight =
            containerWidth / videoRatio;

        offsetX = 0;

        offsetY =
            (containerHeight -
                displayedHeight) / 2;

    } else {

        displayedHeight =
            containerHeight;

        displayedWidth =
            containerHeight * videoRatio;

        offsetY = 0;

        offsetX =
            (containerWidth -
                displayedWidth) / 2;
    }


    const scaleX =
        displayedWidth /
        videoMetadata.width;

    const scaleY =
        displayedHeight /
        videoMetadata.height;


    cropBox.style.left =
        `${offsetX +
          cropState.x * scaleX}px`;

    cropBox.style.top =
        `${offsetY +
          cropState.y * scaleY}px`;

    cropBox.style.width =
        `${cropState.width *
          scaleX}px`;

    cropBox.style.height =
        `${cropState.height *
          scaleY}px`;


    updateInformation();

    updateHiddenInputs();

    updateLivePreview();
}


/* =========================================================
   UPDATE HIDDEN INPUTS
   ========================================================= */

function updateHiddenInputs() {

    if (xInput) {

        xInput.value =
            Math.round(cropState.x);
    }

    if (yInput) {

        yInput.value =
            Math.round(cropState.y);
    }

    if (widthInput) {

        widthInput.value =
            Math.round(cropState.width);
    }

    if (heightInput) {

        heightInput.value =
            Math.round(cropState.height);
    }
}


/* =========================================================
   UPDATE VALUES
   ========================================================= */

function updateInformation() {

    if (cropDimensions) {

        cropDimensions.textContent =
            `${Math.round(cropState.width)} × ` +
            `${Math.round(cropState.height)}`;
    }


    if (cropPosition) {

        cropPosition.textContent =
            `X: ${Math.round(cropState.x)} · ` +
            `Y: ${Math.round(cropState.y)}`;
    }
}


/* =========================================================
   LIVE CROP PREVIEW
   ========================================================= */

function updateLivePreview() {

    if (
        !liveCropPreview ||
        !videoMetadata ||
        videoPreview.readyState < 2
    ) {
        return;
    }


    try {

        const ctx =
            liveCropPreview.getContext("2d");

        if (!ctx) {
            return;
        }


        const width =
            Math.max(
                1,
                Math.round(cropState.width)
            );

        const height =
            Math.max(
                1,
                Math.round(cropState.height)
            );


        liveCropPreview.width = width;

        liveCropPreview.height = height;


        ctx.clearRect(
            0,
            0,
            width,
            height
        );


        ctx.drawImage(

            videoPreview,

            Math.round(cropState.x),

            Math.round(cropState.y),

            width,

            height,

            0,

            0,

            width,

            height
        );


    } catch (error) {

        console.error(
            "Live preview error:",
            error
        );
    }
}


/* =========================================================
   POINTER DOWN
   ========================================================= */

if (cropBox) {

    cropBox.addEventListener(
        "pointerdown",
        function (event) {

            if (
                !videoMetadata ||
                isProcessing
            ) {
                return;
            }


            event.preventDefault();


            const handle =
                event.target.closest(".handle");


            const rect =
                videoEditor.getBoundingClientRect();


            dragState = {

                pointerId:
                    event.pointerId,

                startX:
                    event.clientX,

                startY:
                    event.clientY,

                originalX:
                    cropState.x,

                originalY:
                    cropState.y,

                originalWidth:
                    cropState.width,

                originalHeight:
                    cropState.height,

                handle:
                    handle
                        ? handle.dataset.handle
                        : "move",

                editorWidth:
                    rect.width,

                editorHeight:
                    rect.height
            };


            cropBox.setPointerCapture(
                event.pointerId
            );
        }
    );


    /* =====================================================
       POINTER MOVE
       ===================================================== */

    cropBox.addEventListener(
        "pointermove",
        function (event) {

            if (
                !dragState ||
                !videoMetadata ||
                isProcessing
            ) {
                return;
            }


            event.preventDefault();


            const dx =
                event.clientX -
                dragState.startX;

            const dy =
                event.clientY -
                dragState.startY;


            const rect =
                videoEditor.getBoundingClientRect();


            const videoWidth =
                videoPreview.videoWidth;

            const videoHeight =
                videoPreview.videoHeight;


            const videoRatio =
                videoWidth / videoHeight;

            const containerRatio =
                rect.width / rect.height;


            let displayedWidth;
            let displayedHeight;


            if (videoRatio > containerRatio) {

                displayedWidth =
                    rect.width;

                displayedHeight =
                    rect.width /
                    videoRatio;

            } else {

                displayedHeight =
                    rect.height;

                displayedWidth =
                    rect.height *
                    videoRatio;
            }


            const scaleX =
                videoMetadata.width /
                displayedWidth;

            const scaleY =
                videoMetadata.height /
                displayedHeight;


            const videoDX =
                dx * scaleX;

            const videoDY =
                dy * scaleY;


            const minSize = 40;


            /* =================================================
               MOVE
               ================================================= */

            if (
                dragState.handle === "move"
            ) {

                let newX =
                    dragState.originalX +
                    videoDX;

                let newY =
                    dragState.originalY +
                    videoDY;


                newX =
                    Math.max(
                        0,
                        Math.min(
                            newX,
                            videoMetadata.width -
                            dragState.originalWidth
                        )
                    );


                newY =
                    Math.max(
                        0,
                        Math.min(
                            newY,
                            videoMetadata.height -
                            dragState.originalHeight
                        )
                    );


                cropState.x = newX;
                cropState.y = newY;
            }


            /* =================================================
               NW
               ================================================= */

            if (
                dragState.handle === "nw"
            ) {

                let newX =
                    dragState.originalX +
                    videoDX;

                let newY =
                    dragState.originalY +
                    videoDY;


                newX =
                    Math.max(
                        0,
                        Math.min(
                            newX,
                            dragState.originalX +
                            dragState.originalWidth -
                            minSize
                        )
                    );


                newY =
                    Math.max(
                        0,
                        Math.min(
                            newY,
                            dragState.originalY +
                            dragState.originalHeight -
                            minSize
                        )
                    );


                cropState.x = newX;
                cropState.y = newY;


                cropState.width =
                    dragState.originalWidth -
                    (newX -
                        dragState.originalX);


                cropState.height =
                    dragState.originalHeight -
                    (newY -
                        dragState.originalY);
            }


            /* =================================================
               NE
               ================================================= */

            if (
                dragState.handle === "ne"
            ) {

                let newY =
                    dragState.originalY +
                    videoDY;


                newY =
                    Math.max(
                        0,
                        Math.min(
                            newY,
                            dragState.originalY +
                            dragState.originalHeight -
                            minSize
                        )
                    );


                let newWidth =
                    dragState.originalWidth +
                    videoDX;


                newWidth =
                    Math.max(
                        minSize,
                        Math.min(
                            newWidth,
                            videoMetadata.width -
                            dragState.originalX
                        )
                    );


                cropState.y = newY;

                cropState.width = newWidth;

                cropState.height =
                    dragState.originalHeight -
                    (newY -
                        dragState.originalY);
            }


            /* =================================================
               SW
               ================================================= */

            if (
                dragState.handle === "sw"
            ) {

                let newX =
                    dragState.originalX +
                    videoDX;


                newX =
                    Math.max(
                        0,
                        Math.min(
                            newX,
                            dragState.originalX +
                            dragState.originalWidth -
                            minSize
                        )
                    );


                let newHeight =
                    dragState.originalHeight +
                    videoDY;


                newHeight =
                    Math.max(
                        minSize,
                        Math.min(
                            newHeight,
                            videoMetadata.height -
                            dragState.originalY
                        )
                    );


                cropState.x = newX;

                cropState.width =
                    dragState.originalWidth -
                    (newX -
                        dragState.originalX);

                cropState.height =
                    newHeight;
            }


            /* =================================================
               SE
               ================================================= */

            if (
                dragState.handle === "se"
            ) {

                let newWidth =
                    dragState.originalWidth +
                    videoDX;

                let newHeight =
                    dragState.originalHeight +
                    videoDY;


                newWidth =
                    Math.max(
                        minSize,
                        Math.min(
                            newWidth,
                            videoMetadata.width -
                            dragState.originalX
                        )
                    );


                newHeight =
                    Math.max(
                        minSize,
                        Math.min(
                            newHeight,
                            videoMetadata.height -
                            dragState.originalY
                        )
                    );


                cropState.width =
                    newWidth;

                cropState.height =
                    newHeight;
            }


            updateCropBox();
        }
    );


    /* =====================================================
       POINTER UP
       ===================================================== */

    cropBox.addEventListener(
        "pointerup",
        function () {

            dragState = null;
        }
    );


    cropBox.addEventListener(
        "pointercancel",
        function () {

            dragState = null;
        }
    );
}


/* =========================================================
   VALIDATE CROP
   ========================================================= */

function validateCrop() {

    if (!videoMetadata) {

        showError(
            "Please select a video first."
        );

        return false;
    }


    const x = Number(xInput.value);
    const y = Number(yInput.value);
    const width = Number(widthInput.value);
    const height = Number(heightInput.value);


    if (
        !Number.isFinite(x) ||
        !Number.isFinite(y) ||
        !Number.isFinite(width) ||
        !Number.isFinite(height)
    ) {

        showError(
            "Please select a valid crop area."
        );

        return false;
    }


    if (x < 0 || y < 0) {

        showError(
            "Crop position cannot be negative."
        );

        return false;
    }


    if (width <= 0 || height <= 0) {

        showError(
            "Crop width and height must be greater than zero."
        );

        return false;
    }


    if (
        x + width >
        videoMetadata.width
    ) {

        showError(
            "Crop area exceeds video width."
        );

        return false;
    }


    if (
        y + height >
        videoMetadata.height
    ) {

        showError(
            "Crop area exceeds video height."
        );

        return false;
    }


    return true;
}


/* =========================================================
   FORM SUBMIT
   ========================================================= */

form.addEventListener(
    "submit",
    function (event) {

        /* ---------------------------------------------
           Prevent double submit
           --------------------------------------------- */

        if (isProcessing) {

            event.preventDefault();

            return;
        }


        /* ---------------------------------------------
           Validate crop
           --------------------------------------------- */

        if (!validateCrop()) {

            event.preventDefault();

            return;
        }


        /* ---------------------------------------------
           Ensure latest values are sent
           --------------------------------------------- */

        updateHiddenInputs();


        /* ---------------------------------------------
           Processing state
           --------------------------------------------- */

        isProcessing = true;


        hideOldResult();


        /* ---------------------------------------------
           Hide video
           --------------------------------------------- */

        if (videoEditor) {

            videoEditor.style.display = "none";
        }


        /* ---------------------------------------------
           Hide video column
           --------------------------------------------- */

        const videoColumn =
            document.getElementById("videoColumn");


        if (videoColumn) {

            videoColumn.style.display = "none";
        }


        /* ---------------------------------------------
           Expand side panel
           --------------------------------------------- */

        const sidePanel =
            document.getElementById("sidePanel");


        if (sidePanel) {

            sidePanel.style.width = "100%";
        }


        /* ---------------------------------------------
           Disable process button
           --------------------------------------------- */

        if (processButton) {

            processButton.disabled = true;

            processButton.textContent =
                "Processing...";
        }


        /* ---------------------------------------------
           Show loading
           --------------------------------------------- */

        if (loading) {

            loading.style.display = "block";
        }

    }
);


/* =========================================================
   RESPONSIVE UPDATE
   ========================================================= */

window.addEventListener(
    "resize",
    function () {

        if (videoMetadata) {

            updateCropBox();
        }
    }
);


/* =========================================================
   CLEANUP
   ========================================================= */

window.addEventListener(
    "beforeunload",
    function () {

        if (objectURL) {

            URL.revokeObjectURL(objectURL);

            objectURL = null;
        }


        clearTimeout(popupTimer);
        clearTimeout(popupHideTimer);
    }
);


/* =========================================================
   INITIALIZATION
   ========================================================= */

console.log(
    "Crop Video Editor initialized."
);


});
