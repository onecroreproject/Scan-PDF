document.addEventListener('DOMContentLoaded', () => {
    const pricing = {
        monthly: {
            free: 0,
            pro: 499,
            business: 999,
            businessPlus: 1999
        },
        yearly: {
            free: 0,
            pro: 4999,
            business: 9999,
            businessPlus: 19999
        }
    };

    const toggleBtn = document.getElementById('billing-toggle');
    const toggleCircle = document.getElementById('toggle-circle');
    const monthlyLabel = document.getElementById('billing-monthly-label');
    const yearlyLabel = document.getElementById('billing-yearly-label');
    
    const cards = {
        free: {
            priceEl: document.getElementById('price-free'),
            periodEl: document.getElementById('period-free'),
            key: 'free'
        },
        pro: {
            priceEl: document.getElementById('price-pro'),
            periodEl: document.getElementById('period-pro'),
            key: 'pro'
        },
        business: {
            priceEl: document.getElementById('price-business'),
            periodEl: document.getElementById('period-business'),
            key: 'business'
        },
        businessPlus: {
            priceEl: document.getElementById('price-business-plus'),
            periodEl: document.getElementById('period-business-plus'),
            key: 'businessPlus'
        }
    };

    let isYearly = false;

    // Currency Formatting using Intl.NumberFormat('en-IN')
    const formatINR = (value) => {
        return '₹' + new Intl.NumberFormat('en-IN').format(value);
    };

    const updatePricing = () => {
        const mode = isYearly ? 'yearly' : 'monthly';
        const periodText = isYearly ? '/year' : '/month';

        // Update payment hrefs dynamically
        const planButtons = document.querySelectorAll('.plan-btn');
        planButtons.forEach(btn => {
            const planCode = btn.getAttribute('data-plan');
            if (planCode) {
                btn.setAttribute('href', `/services/payment/confirm/${planCode}/${mode}/`);
            }
        });

        Object.values(cards).forEach(card => {
            if (!card.priceEl) return;

            // Fade Out (150ms)
            card.priceEl.classList.add('price-fade-out');
            if (card.periodEl) {
                card.periodEl.classList.add('price-fade-out');
            }

            setTimeout(() => {
                const value = pricing[mode][card.key];
                card.priceEl.textContent = formatINR(value);
                if (card.periodEl) {
                    card.periodEl.textContent = periodText;
                }

                // Switch classes to Fade In (150ms)
                card.priceEl.classList.remove('price-fade-out');
                card.priceEl.classList.add('price-fade-in');
                if (card.periodEl) {
                    card.periodEl.classList.remove('price-fade-out');
                    card.periodEl.classList.add('price-fade-in');
                }

                // Cleanup animation classes
                setTimeout(() => {
                    card.priceEl.classList.remove('price-fade-in');
                    if (card.periodEl) {
                        card.periodEl.classList.remove('price-fade-in');
                    }
                }, 150);
            }, 150);
        });
    };

    if (toggleBtn) {
        toggleBtn.addEventListener('click', () => {
            isYearly = !isYearly;

            if (isYearly) {
                if (toggleCircle) toggleCircle.style.transform = 'translateX(1.75rem)';
                if (monthlyLabel && yearlyLabel) {
                    monthlyLabel.classList.replace('text-brand-600', 'text-surface-400');
                    yearlyLabel.classList.replace('text-surface-400', 'text-brand-600');
                }
            } else {
                if (toggleCircle) toggleCircle.style.transform = 'translateX(0)';
                if (monthlyLabel && yearlyLabel) {
                    yearlyLabel.classList.replace('text-brand-600', 'text-surface-400');
                    monthlyLabel.classList.replace('text-surface-400', 'text-brand-600');
                }
            }

            updatePricing();
        });
    }

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
