/* ═══════════════════════════════════════════════════════════
   ConvertPro - Main JavaScript
   Handles navbar behavior and mobile menu
   ═══════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', () => {
    // ─── Navbar Scroll Effect ─────────────────────────────
    const navbar = document.getElementById('navbar');
    
    const handleScroll = () => {
        if (window.scrollY > 50) {
            navbar.classList.add('scrolled');
        } else {
            navbar.classList.remove('scrolled');
        }
    };
    
    window.addEventListener('scroll', handleScroll, { passive: true });
    handleScroll(); // Initial check
    
    
    // ─── Mobile Menu Toggle ───────────────────────────────
    const mobileMenuBtn = document.getElementById('mobile-menu-btn');
    const mobileMenu = document.getElementById('mobile-menu');
    
    if (mobileMenuBtn && mobileMenu) {
        mobileMenuBtn.addEventListener('click', () => {
            const isOpen = !mobileMenu.classList.contains('hidden');
            
            if (isOpen) {
                mobileMenu.classList.add('hidden');
                mobileMenuBtn.innerHTML = '<i data-lucide="menu" class="w-6 h-6"></i>';
            } else {
                mobileMenu.classList.remove('hidden');
                mobileMenuBtn.innerHTML = '<i data-lucide="x" class="w-6 h-6"></i>';
            }
            
            lucide.createIcons();
        });
        
        // Close menu when a link is clicked
        mobileMenu.querySelectorAll('a').forEach(link => {
            link.addEventListener('click', () => {
                mobileMenu.classList.add('hidden');
                mobileMenuBtn.innerHTML = '<i data-lucide="menu" class="w-6 h-6"></i>';
                lucide.createIcons();
            });
        });
    }
    
    
    // ─── Smooth Scroll for Anchor Links ───────────────────
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', (e) => {
            const targetId = anchor.getAttribute('href');
            if (targetId === '#') return;
            
            const target = document.querySelector(targetId);
            if (target) {
                e.preventDefault();
                const offset = navbar.offsetHeight + 20;
                const top = target.getBoundingClientRect().top + window.scrollY - offset;
                
                window.scrollTo({
                    top: top,
                    behavior: 'smooth'
                });
            }
        });
    });
    
    
    // ─── Intersection Observer for Scroll Animations ──────
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    };
    
    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
                observer.unobserve(entry.target);
            }
        });
    }, observerOptions);
    
    // Observe tool cards for scroll-triggered animations
    document.querySelectorAll('.tool-card').forEach(card => {
        card.style.opacity = '0';
        card.style.transform = 'translateY(20px)';
        card.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(card);
    });
});
