document.addEventListener("DOMContentLoaded", function() {
    // Current Date and Time in Top Navbar
    const datetimeElement = document.getElementById('current-datetime');
    if (datetimeElement) {
        setInterval(() => {
            const now = new Date();
            const options = { weekday: 'long', year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' };
            datetimeElement.textContent = now.toLocaleDateString('en-US', options);
        }, 1000);
    }
});
