/* ═══════════════════════════════════════════════════════════
   ConvertPro - Main JavaScript
   Handles navbar behavior and mobile menu
   ═══════════════════════════════════════════════════════════ */

(function initializeScanPDFNavbar() {
    if (window.__scanpdfNavbarInitialized) {
        return;
    }

    window.__scanpdfNavbarInitialized = true;

    const init = () => {
        const navbar = document.getElementById('navbar');

        const handleScroll = () => {
            if (!navbar) return;
            if (window.scrollY > 50) {
                navbar.classList.add('scrolled');
            } else {
                navbar.classList.remove('scrolled');
            }
        };

        if (navbar) {
            window.addEventListener('scroll', handleScroll, { passive: true });
            handleScroll();
        }

        const mobileMenuBtn = document.getElementById('mobile-menu-btn');
        const mobileDrawer = document.getElementById('mobile-drawer');
        const mobileDrawerOverlay = document.getElementById('mobile-drawer-overlay');
        const closeDrawerBtn = document.getElementById('close-drawer-btn');

        const setDrawerState = (isOpen) => {
            if (!mobileMenuBtn || !mobileDrawer || !mobileDrawerOverlay) return;

            mobileDrawer.style.transform = isOpen ? 'translateX(0)' : 'translateX(100%)';
            mobileDrawerOverlay.style.opacity = isOpen ? '1' : '0';
            mobileDrawerOverlay.style.visibility = isOpen ? 'visible' : 'hidden';
            mobileMenuBtn.setAttribute('aria-expanded', String(isOpen));
            mobileMenuBtn.innerHTML = isOpen
                ? '<i data-lucide="x" class="w-6 h-6"></i>'
                : '<i data-lucide="menu" class="w-6 h-6"></i>';

            document.body.style.overflow = isOpen ? 'hidden' : '';

            if (typeof lucide !== 'undefined') {
                lucide.createIcons();
            }

            // Reset accordions when closing drawer
            if (!isOpen) {
                const accordions = document.querySelectorAll('.accordion-btn');
                accordions.forEach(btn => {
                    btn.setAttribute('aria-expanded', 'false');
                    const content = btn.nextElementSibling;
                    if (content) content.classList.add('hidden');
                    const chevron = btn.querySelector('[data-lucide="chevron-down"], [data-lucide="chevron-right"]');
                    if (chevron) chevron.style.transform = '';
                });
            }
        };

        if (mobileMenuBtn && mobileDrawer && mobileDrawerOverlay) {
            mobileMenuBtn.addEventListener('click', () => {
                const isOpen = mobileMenuBtn.getAttribute('aria-expanded') === 'true';
                setDrawerState(!isOpen);
            });

            if (closeDrawerBtn) {
                closeDrawerBtn.addEventListener('click', () => setDrawerState(false));
            }

            mobileDrawerOverlay.addEventListener('click', () => setDrawerState(false));

            mobileDrawer.querySelectorAll('a').forEach((link) => {
                link.addEventListener('click', () => setDrawerState(false));
            });

            document.addEventListener('keydown', (event) => {
                if (event.key === 'Escape') {
                    setDrawerState(false);
                }
            });
        }

        document.querySelectorAll('a[href^="#"]').forEach((anchor) => {
            anchor.addEventListener('click', (e) => {
                const targetId = anchor.getAttribute('href');
                if (targetId === '#') return;

                const target = document.querySelector(targetId);
                if (!target || !navbar) {
                    return;
                }

                e.preventDefault();
                const offset = navbar.offsetHeight + 20;
                const top = target.getBoundingClientRect().top + window.scrollY - offset;

                window.scrollTo({
                    top: top,
                    behavior: 'smooth'
                });
            });
        });

        if ('IntersectionObserver' in window) {
            const observerOptions = {
                threshold: 0.1,
                rootMargin: '0px 0px -50px 0px'
            };

            const observer = new IntersectionObserver((entries) => {
                entries.forEach((entry) => {
                    if (entry.isIntersecting) {
                        entry.target.style.opacity = '1';
                        entry.target.style.transform = 'translateY(0)';
                        observer.unobserve(entry.target);
                    }
                });
            }, observerOptions);

            document.querySelectorAll('.tool-card').forEach((card) => {
                card.style.opacity = '0';
                card.style.transform = 'translateY(20px)';
                card.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
                observer.observe(card);
            });
        }
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init, { once: true });
    } else {
        init();
    }
})();
