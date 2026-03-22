(function () {
    'use strict';

    // ── Config ──
    const TOTAL_SLIDES = 10;
    const TRANSITION_DURATION = 1.2;
    const ORBIT_DURATION = 10;

    // ── State ──
    let currentSlide = 0;
    let isTransitioning = false;
    let orbitAnim = null;

    // ── DOM refs ──
    let slidesContainer, batteryFill, batteryText,
        currentSlideEl, navHint, loader, carWrapper;

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
        navHint         = document.getElementById('navHint');
        loader          = document.getElementById('loader');
        carWrapper      = document.getElementById('carWrapper');

        buildTrail();
        initCharts();
        initEvents();
        updateBattery();
        animateSlideContent(0);
        startOrbit();

        // hide loader
        setTimeout(() => {
            loader.classList.add('hidden');
            document.querySelector('.battery-indicator').classList.add('visible');
            document.querySelector('.slide-counter').classList.add('visible');
        }, 1200);

        // hide nav hint
        setTimeout(() => {
            gsap.to(navHint, { opacity: 0, duration: 0.8, onComplete: () => navHint.style.display = 'none' });
        }, 5000);
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
    //  ORBIT (slide 1 only)
    // =============================================
    function startOrbit() {
        if (orbitAnim) orbitAnim.kill();
        clearTrail();
        carWrapper.classList.remove('hidden');

        const vw = window.innerWidth;
        const vh = window.innerHeight;
        const cx = vw / 2, cy = vh / 2;
        const rx = Math.min(vw * 0.32, 420);
        const ry = Math.min(vh * 0.30, 230);
        const obj = { a: 0 };

        orbitAnim = gsap.to(obj, {
            a: Math.PI * 2,
            duration: ORBIT_DURATION,
            repeat: -1,
            ease: 'none',
            onUpdate: () => {
                const x = cx + Math.cos(obj.a) * rx;
                const y = cy + Math.sin(obj.a) * ry;
                const dx = -Math.sin(obj.a) * rx;
                const dy =  Math.cos(obj.a) * ry;
                const angle = Math.atan2(dx, dy) * (180 / Math.PI);

                carWrapper.style.left = x + 'px';
                carWrapper.style.top = y + 'px';
                carWrapper.style.transform = 'translate(-50%, -50%) rotate(' + (-angle) + 'deg) scale(1)';
                updateTrail();
            }
        });
    }

    function stopOrbit() {
        if (orbitAnim) { orbitAnim.kill(); orbitAnim = null; }
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
        const goingToTitle = (idx === 0);
        const leavingTitle = (currentSlide === 0);

        if (leavingTitle) stopOrbit();
        clearTrail();

        // Show car for transition animation
        carWrapper.classList.remove('hidden');

        const tl = gsap.timeline({
            onComplete: () => {
                currentSlide = idx;
                isTransitioning = false;
                updateBattery();
                updateCounter();

                if (goingToTitle) {
                    startOrbit();
                } else {
                    // Hide car on content slides
                    carWrapper.classList.add('hidden');
                }

                // Replay charts when their slides appear
                if (currentSlide === 2) replayChart(salesChart);
                if (currentSlide === 4) replayChart(demoChart);
                animateSlideContent(currentSlide);
            }
        });

        // If not coming from title, set car at edge for drive-through
        if (!leavingTitle) {
            gsap.set(carWrapper, {
                left: dir > 0 ? -80 : vw + 80,
                top: vh / 2,
                rotation: dir > 0 ? -90 : 90,
                scale: 0.8
            });
        }

        // 1. Car rotates to face travel direction
        if (leavingTitle) {
            tl.to(carWrapper, {
                rotation: dir > 0 ? -90 : 90,
                duration: 0.2,
                ease: 'power2.inOut'
            }, 0);
        }

        // 2. Car drives across screen
        const carDriveStart = leavingTitle ? 0.1 : 0;
        tl.to(carWrapper, {
            left: dir > 0 ? vw + 100 : -100,
            top: vh / 2,
            duration: 0.6,
            ease: 'power2.inOut'
        }, carDriveStart);

        tl.to(carWrapper, {
            scale: 1.1,
            duration: 0.3,
            ease: 'power2.in'
        }, carDriveStart);

        // 3. Slides translate
        tl.to(slidesContainer, {
            x: -idx * vw,
            duration: TRANSITION_DURATION,
            ease: 'power3.inOut'
        }, 0.08);

        // 4. If going to title, car reappears and starts orbiting (handled in onComplete)
        if (goingToTitle) {
            // Set car offscreen, orbit will position it
            tl.set(carWrapper, {
                left: vw / 2,
                top: vh / 2,
                rotation: 0,
                scale: 1
            }, TRANSITION_DURATION - 0.1);
        }

        // 5. Parallax on incoming slide content
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
            '.slide-heading, .section-number, .slide-subtitle, ' +
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
    //  CHARTS
    // =============================================
    let salesChart, demoChart;

    function initCharts() {
        const red = 'rgba(232,33,39,';
        const sCtx = document.getElementById('salesChart');
        if (sCtx) {
            salesChart = new Chart(sCtx, {
                type: 'bar',
                data: {
                    labels: ['2019', '2020', '2021', '2022', '2023', '2024', '2025'],
                    datasets: [{
                        label: 'Model 3 US Sales (thousands)',
                        data: [154, 122, 176, 252, 209, 235, 268],
                        backgroundColor: [
                            red + '0.25)', red + '0.3)', red + '0.4)',
                            red + '0.6)', red + '0.5)', red + '0.75)', red + '0.9)'
                        ],
                        borderColor: red + '1)',
                        borderWidth: 2,
                        borderRadius: 8
                    }]
                },
                options: chartOpts(true)
            });
        }

        const dCtx = document.getElementById('demoChart');
        if (dCtx) {
            demoChart = new Chart(dCtx, {
                type: 'doughnut',
                data: {
                    labels: ['Tech Pros', 'Eco-Conscious', 'Early Adopters', 'Luxury Buyers', 'Fleet Ops'],
                    datasets: [{
                        data: [35, 25, 20, 12, 8],
                        backgroundColor: ['#e82127', '#ff5722', '#ff9800', '#ffc107', '#ffeb3b'],
                        borderWidth: 0,
                        hoverOffset: 10
                    }]
                },
                options: chartOpts(false)
            });
        }
    }

    function chartOpts(isBar) {
        const base = {
            responsive: true,
            maintainAspectRatio: false,
            animation: { duration: 1400, easing: 'easeOutQuart' },
            plugins: {
                legend: {
                    labels: {
                        font: { family: "'Inter', sans-serif", size: 12 },
                        usePointStyle: true,
                        padding: 14,
                        color: 'rgba(255,255,255,0.7)'
                    }
                }
            }
        };
        if (isBar) {
            base.scales = {
                y: {
                    beginAtZero: true,
                    grid: { color: 'rgba(255,255,255,0.06)' },
                    ticks: { font: { family: "'Inter'" }, color: 'rgba(255,255,255,0.5)' }
                },
                x: {
                    grid: { display: false },
                    ticks: { font: { family: "'Inter'" }, color: 'rgba(255,255,255,0.5)' }
                }
            };
            base.plugins.legend.position = 'top';
        } else {
            base.plugins.legend.position = 'bottom';
        }
        return base;
    }

    function replayChart(chart) {
        if (!chart) return;
        chart.reset();
        chart.update();
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
