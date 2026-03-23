(function () {
    'use strict';

    // ── Config ──
    const TOTAL_SLIDES = 7;
    const TRANSITION_DURATION = 1.2;
    const ORBIT_DURATION = 10;

    // ── State ──
    let currentSlide = 0;
    let isTransitioning = false;
    let orbitAnim = null;

    // ── DOM refs ──
    let slidesContainer, batteryFill, batteryText,
        currentSlideEl, loader, carWrapper;

    // ── Trail ──
    let trailPositions = [];
    const TRAIL_COUNT = 22;
    let trailEls = [];

    // =============================================
    //  INIT
    // =============================================
    function init() {
        slidesContainer = document.getElementById('slidesContainer');
        batteryFill     = document.getElementById('batteryFill');
        batteryText     = document.getElementById('batteryText');
        currentSlideEl  = document.getElementById('currentSlide');
        loader          = document.getElementById('loader');
        carWrapper      = document.getElementById('carWrapper');

        buildTrail();
        buildSlideNav();
        initEvents();
        updateBattery();
        updateSlideNav();
        animateSlideContent(0);
        startOrbit();

        // hide loader
        setTimeout(() => {
            loader.classList.add('hidden');
            document.querySelector('.battery-indicator').classList.add('visible');
            document.querySelector('.slide-counter').classList.add('visible');
            document.querySelector('.slide-nav').classList.add('visible');
        }, 1200);

    }

    // =============================================
    //  TRAIL
    // =============================================
    function buildTrail() {
        const container = document.getElementById('carTrail');
        for (let i = 0; i < TRAIL_COUNT; i++) {
            const t = 1 - i / TRAIL_COUNT;
            const dot = document.createElement('div');
            dot.className = 'trail-dot';
            dot.style.opacity = 0.3 * t;
            dot.style.width = dot.style.height = (5 * t) + 'px';
            dot.style.display = 'none';
            container.appendChild(dot);
            trailEls.push(dot);
        }
    }

    function updateTrail() {
        const rect = carWrapper.getBoundingClientRect();
        const cx = rect.left + rect.width / 2;
        const cy = rect.top + rect.height / 2;
        trailPositions.unshift({ x: cx, y: cy });
        if (trailPositions.length > TRAIL_COUNT) trailPositions.pop();
        trailEls.forEach((dot, i) => {
            if (i < trailPositions.length) {
                dot.style.display = 'block';
                dot.style.left = trailPositions[i].x + 'px';
                dot.style.top = trailPositions[i].y + 'px';
            } else {
                dot.style.display = 'none';
            }
        });
    }

    function clearTrail() {
        trailPositions = [];
        trailEls.forEach(d => d.style.display = 'none');
    }

    // =============================================
    //  FREE ROAM (slide 1 only)
    // =============================================
    let carX = 0, carY = 0, carAngle = -90; // angle in degrees, -90 = facing up
    let carSpeed = 2.5;
    let roamRAF = null;
    let turnTimer = null;

    function startOrbit() {
        stopOrbit();
        clearTrail();
        carWrapper.classList.remove('hidden');

        const vw = window.innerWidth;
        const vh = window.innerHeight;
        carX = vw / 2;
        carY = vh / 2;
        carAngle = -90 + (Math.random() * 60 - 30); // start roughly upward
        carSpeed = 2.5;

        // Random turn every 3-6 seconds
        function scheduleTurn() {
            const delay = 3000 + Math.random() * 3000;
            turnTimer = setTimeout(() => {
                const turnAmount = (Math.random() > 0.5 ? 1 : -1) * (40 + Math.random() * 50);
                gsap.to({ val: carAngle }, {
                    val: carAngle + turnAmount,
                    duration: 0.8,
                    ease: 'power2.inOut',
                    onUpdate: function() { carAngle = this.targets()[0].val; }
                });
                scheduleTurn();
            }, delay);
        }
        scheduleTurn();

        function tick() {
            const rad = carAngle * Math.PI / 180;
            carX += Math.cos(rad) * carSpeed;
            carY += Math.sin(rad) * carSpeed;

            const vw = window.innerWidth;
            const vh = window.innerHeight;
            const margin = 60;

            // Wrap around edges
            if (carX < -margin) carX = vw + margin;
            if (carX > vw + margin) carX = -margin;
            if (carY < -margin) carY = vh + margin;
            if (carY > vh + margin) carY = -margin;

            carWrapper.style.left = carX + 'px';
            carWrapper.style.top = carY + 'px';
            // SVG car points up (0deg), so rotate by carAngle + 90
            carWrapper.style.transform = 'translate(-50%, -50%) rotate(' + (carAngle + 90) + 'deg) scale(1)';

            updateTrail();
            roamRAF = requestAnimationFrame(tick);
        }
        roamRAF = requestAnimationFrame(tick);
    }

    function stopOrbit() {
        if (roamRAF) { cancelAnimationFrame(roamRAF); roamRAF = null; }
        if (turnTimer) { clearTimeout(turnTimer); turnTimer = null; }
    }

    // =============================================
    //  NAVIGATION
    // =============================================
    function goTo(idx) {
        if (isTransitioning) return;
        if (idx < 0 || idx >= TOTAL_SLIDES || idx === currentSlide) return;
        isTransitioning = true;

        const dir = idx > currentSlide ? 1 : -1;
        const vw = window.innerWidth, vh = window.innerHeight;

        // Pause roaming during transition
        stopOrbit();
        clearTrail();
        carWrapper.classList.remove('hidden');

        const tl = gsap.timeline({
            onComplete: () => {
                currentSlide = idx;
                isTransitioning = false;
                updateBattery();
                updateCounter();
                updateSlideNav();

                // Resume roaming on every slide
                startOrbit();

                animateSlideContent(currentSlide);
            }
        });

        // 1. Slides translate
        tl.to(slidesContainer, {
            x: -idx * vw,
            duration: TRANSITION_DURATION,
            ease: 'power3.inOut'
        }, 0);

        // 2. Parallax on incoming slide content
        const nextSlideEl = document.querySelectorAll('.slide')[idx];
        if (nextSlideEl) {
            const content = nextSlideEl.querySelector('.slide-content');
            if (content) {
                gsap.fromTo(content,
                    { x: dir * 80 },
                    { x: 0, duration: TRANSITION_DURATION, ease: 'power3.inOut', delay: 0.06 }
                );
            }
        }
    }

    // =============================================
    //  SLIDE NAV BAR
    // =============================================
    function buildSlideNav() {
        const nav = document.getElementById('slideNav');
        for (let i = 0; i < TOTAL_SLIDES; i++) {
            const item = document.createElement('div');
            item.className = 'slide-nav-item';
            item.textContent = String(i + 1).padStart(2, '0');
            item.addEventListener('click', () => goTo(i));
            nav.appendChild(item);
        }
    }

    function updateSlideNav() {
        const items = document.querySelectorAll('.slide-nav-item');
        items.forEach((item, i) => {
            item.classList.toggle('active', i === currentSlide);
        });
    }

    // =============================================
    //  UI UPDATES
    // =============================================
    function updateBattery() {
        const pct = Math.round(100 - (currentSlide / (TOTAL_SLIDES - 1)) * 100);
        batteryFill.style.width = pct + '%';
        batteryText.textContent = pct + '%';
        if (pct > 60)      batteryFill.style.background = 'linear-gradient(90deg, #4caf50, #8bc34a)';
        else if (pct > 30) batteryFill.style.background = 'linear-gradient(90deg, #ff9800, #ffc107)';
        else               batteryFill.style.background = 'linear-gradient(90deg, #e82127, #ff5722)';
    }

    function updateCounter() {
        currentSlideEl.textContent = String(currentSlide + 1).padStart(2, '0');
    }

    function animateSlideContent(idx) {
        const slide = document.querySelectorAll('.slide')[idx];
        if (!slide) return;
        const els = slide.querySelectorAll(
            '.info-card, .strategy-card, .takeaway, .demo-stat, ' +
            '.slide-heading, .section-number, .slide-label, .slide-subtitle, ' +
            '.chart-wrap, .demo-chart-wrap, .closing-line, ' +
            '.hero-line, .overline, .hero-tagline, .title-divider, ' +
            '.compare-col, .quote-block, .metric, .timeline-item, ' +
            '.pricing-hero, .pill'
        );

        // Animate slide-up only, keep elements always visible
        els.forEach((el, i) => {
            el.style.opacity = '1';
            gsap.fromTo(el,
                { y: 18 },
                { y: 0, duration: 0.45, delay: i * 0.05, ease: 'power2.out' }
            );
        });
    }

    // =============================================
    //  EVENTS
    // =============================================
    function initEvents() {
        // Keyboard
        document.addEventListener('keydown', e => {
            if (e.key === 'ArrowRight' || e.key === 'ArrowDown') { e.preventDefault(); goTo(currentSlide + 1); }
            if (e.key === 'ArrowLeft'  || e.key === 'ArrowUp')   { e.preventDefault(); goTo(currentSlide - 1); }
        });

        // Scroll wheel
        let wt;
        document.addEventListener('wheel', e => {
            e.preventDefault();
            clearTimeout(wt);
            wt = setTimeout(() => {
                (e.deltaY > 0 || e.deltaX > 0) ? goTo(currentSlide + 1) : goTo(currentSlide - 1);
            }, 60);
        }, { passive: false });

        // Touch swipe
        let tx = 0;
        document.addEventListener('touchstart', e => { tx = e.touches[0].clientX; });
        document.addEventListener('touchend', e => {
            const d = tx - e.changedTouches[0].clientX;
            if (Math.abs(d) > 50) d > 0 ? goTo(currentSlide + 1) : goTo(currentSlide - 1);
        });

        // Click zones on edges
        document.getElementById('clickLeft').addEventListener('click', () => goTo(currentSlide - 1));
        document.getElementById('clickRight').addEventListener('click', () => goTo(currentSlide + 1));

        // Resize handler
        window.addEventListener('resize', () => {
            gsap.set(slidesContainer, { x: -currentSlide * window.innerWidth });
            if (currentSlide === 0 && orbitAnim) {
                stopOrbit();
                startOrbit();
            }
        });
    }

    // =============================================
    //  START
    // =============================================
    document.addEventListener('DOMContentLoaded', init);
})();
