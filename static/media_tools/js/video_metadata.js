function loadVideoMetadata(
    file,
    callback
) {

    try {

        if (!file) {
            return;
        }

        const video =
            document.createElement(
                "video"
            );

        video.preload = "metadata";

        video.onloadedmetadata =
            function () {

                const metadata = {

                    duration:
                        video.duration,

                    width:
                        video.videoWidth,

                    height:
                        video.videoHeight

                };

                URL.revokeObjectURL(
                    video.src
                );

                callback(metadata);
            };

        video.onerror =
            function () {

                showValidationError(
                    "Unable to read this video."
                );

            };

        video.src =
            URL.createObjectURL(file);

    } catch (error) {

        console.error(
            "Metadata error:",
            error
        );

        showValidationError(
            "Unable to read video information."
        );
    }
}