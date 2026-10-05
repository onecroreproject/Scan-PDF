document.addEventListener('DOMContentLoaded', () => {
    // Note: Monthly/Yearly toggle logic has been moved to pricing.html 
    // to dynamically read prices and feature limits from data attributes.


    // Modal upgrade handlers bound to window object
    window.openUpgradeModal = function(planName) {
        const modal = document.getElementById('upgrade-modal');
        const modalContent = modal ? modal.firstElementChild : null;
        const planNameSpan = document.getElementById('modal-plan-name');

        if (modal && modalContent && planNameSpan) {
            planNameSpan.textContent = planName;
            modal.classList.remove('opacity-0', 'pointer-events-none');
            modalContent.classList.remove('scale-95');
            modalContent.classList.add('scale-100');
        }
    };

    window.closeUpgradeModal = function() {
        const modal = document.getElementById('upgrade-modal');
        const modalContent = modal ? modal.firstElementChild : null;

        if (modal && modalContent) {
            modal.classList.add('opacity-0', 'pointer-events-none');
            modalContent.classList.remove('scale-100');
            modalContent.classList.add('scale-95');
        }
    };
});
