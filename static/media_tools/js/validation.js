function showValidationError(message) {

    try {

        const popup =
            document.getElementById(
                "validationPopup"
            );

        if (!popup) {
            return;
        }

        popup.textContent = message;

        popup.classList.add("show");

        clearTimeout(
            window.validationTimer
        );

        window.validationTimer =
            setTimeout(
                function () {

                    popup.classList.remove(
                        "show"
                    );

                },
                2500
            );

    } catch (error) {

        console.error(
            "Validation popup error:",
            error
        );
    }
}