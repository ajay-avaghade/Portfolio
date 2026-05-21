// Initialize Lucide Icons
lucide.createIcons();

// Intersection Observer for scroll animations
document.addEventListener("DOMContentLoaded", () => {
    const observerOptions = {
        root: null,
        rootMargin: '0px',
        threshold: 0.1
    };

    const observer = new IntersectionObserver((entries, observer) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('is-visible');
                // Optional: Stop observing once animated
                // observer.unobserve(entry.target);
            }
        });
    }, observerOptions);

    const animatedElements = document.querySelectorAll('.animate-on-scroll');
    
    // Trigger animations for elements already in viewport on load, delay others
    animatedElements.forEach((el, index) => {
        observer.observe(el);
    });

    // ------------------------------------------------
    // GAMIFIED CHECKOUT FLOW LOGIC
    // ------------------------------------------------

    // 1. Sticky Footer Visibility based on Scroll
    const stickyCart = document.getElementById('sticky-cart');
    const heroSection = document.getElementById('about');
    const checkoutSection = document.getElementById('checkout');

    window.addEventListener('scroll', () => {
        const scrollPos = window.scrollY;
        const heroBottom = heroSection.getBoundingClientRect().bottom + window.scrollY;
        const checkoutTop = checkoutSection.getBoundingClientRect().top + window.scrollY;

        // Show sticky cart ONLY between Hero and Checkout sections
        if (scrollPos > heroBottom && scrollPos < (checkoutTop - window.innerHeight + 100)) {
            stickyCart.classList.add('visible');
        } else {
            stickyCart.classList.remove('visible');
        }
    });

    // 2. Promo Code "Apply" Logic
    const applyBtn = document.getElementById('apply-promo-btn');
    const successText = document.getElementById('promo-success');
    const checkoutBtn = document.getElementById('final-checkout-btn');

    if(applyBtn) {
        applyBtn.addEventListener('click', () => {
            // Disable apply button to prevent spamming
            applyBtn.textContent = 'Applied';
            applyBtn.disabled = true;
            applyBtn.style.color = 'var(--success)';
            applyBtn.style.borderColor = 'var(--success)';

            // Show success text
            successText.style.display = 'flex';

            // Enable final checkout button (Unlock gamified progression)
            checkoutBtn.classList.remove('disabled');
            
            // Optionally, we can also style the total value to show a struck-through old CTC
            // but for simplicity, we just enable the button.
            
            // Small pop animation for the checkout button
            checkoutBtn.style.transform = 'scale(1.05)';
            setTimeout(() => {
                checkoutBtn.style.transform = 'scale(1)';
            }, 200);
        });
    }

    // ------------------------------------------------
    // DYNAMIC ATTENTION-HOOK ANIMATIONS
    // ------------------------------------------------

    // 1. Typewriter Effect
    const typeTarget = document.getElementById('typewriter');
    const typeText = "Data-Driven Product Manager";
    let charIndex = 0;
    
    // Tiny delay before starting to let page render
    setTimeout(() => {
        function typeWriter() {
            if (charIndex < typeText.length) {
                typeTarget.innerHTML += typeText.charAt(charIndex);
                charIndex++;
                setTimeout(typeWriter, 50); // Speed of typing
            }
        }
        if(typeTarget) typeWriter();
    }, 500);

    // 2. Glow-on-Hover Spotlight Effect
    const glowCards = document.querySelectorAll('.metric-card, .case-card, .invoice-card');
    
    glowCards.forEach(card => {
        card.addEventListener('mousemove', e => {
            const rect = card.getBoundingClientRect();
            // Calculate mouse position strictly relative to the card's bounding box
            const x = e.clientX - rect.left;
            const y = e.clientY - rect.top;
            
            // Set css variables local to this card
            card.style.setProperty('--mouse-x', `${x}px`);
            card.style.setProperty('--mouse-y', `${y}px`);
        });
    });
    // ------------------------------------------------
    // CASE STUDY MODAL LOGIC
    // ------------------------------------------------
    const caseModal = document.getElementById('case-modal');
    const modalBody = document.getElementById('modal-body');
    const closeBtn = document.getElementById('close-modal');
    const readMoreBtns = document.querySelectorAll('.read-more');

    const caseStudies = {
        'offers-revamp': {
            category: 'Director Deep-Dive @ PhonePe',
            title: 'Offer Discovery & Checkout Revamp',
            hook: 'Countering market share attrition through design-led incentivization and algorithmic efficiency.',
            narrative: [
                {
                    type: 'intro',
                    title: 'The Invisible Offer Crisis',
                    text: 'Despite a robust catalog, PhonePe lost 300 bps market share (49% to 46%) to entrants like Navi and Super Money who flooded the market with cashbacks. User research revealed a stark disconnect: users weren\'t churning because PhonePe lacked offers; they were churning because the offers were "Invisible."',
                    voc: [
                        '“PhonePe isn’t giving offers anymore.”',
                        '“I don’t know where to find recharges/bill payment offers.”',
                        '“I can’t find the reward I was promised.”',
                        '“No offers on my preferred payment instruments.”'
                    ]
                },
                {
                    type: 'strategy',
                    title: 'Strategic Pivot: Propensity-to-Transact',
                    text: 'Instead of a blind price war, we identified "switchers"—price-sensitive users migrating for rewards. I architected a \'propensity-to-transact\' model to intelligently cohort users, targeting aggressive offers only toward high-flight-risk seekers, reducing overall marketing burn by 32%.'
                },
                {
                    type: 'design',
                    title: 'Solution 1: Contextual Discovery',
                    text: 'Brought discovery out of hidden menus directly onto the homepage category tiles and integrated with the recommendation engine to showcase relevant offers at high-intent placements throughout the app.',
                    image: 'assets/discovery.jpg',
                    caption: 'Homepage integration bringing visibility to category-level incentives.'
                },
                {
                    type: 'figma_breakdown',
                    title: 'Solution 2: The "Stacks & Sections" Checkout',
                    text: 'We overhauled the checkout logically to handle complex incentivization. Instead of showing only the auto-applied offer, we introduced "Supplementary Constraints" to give users visibility into all probable rewards.',
                    images: [
                        { url: 'assets/checkout_multi_offers.png', caption: 'Figma Breakdown: Users see multiple offers per instrument (e.g., RuPay Credit Card showing ₹100 + 10% cashback), giving them clear agency and choice.' },
                        { url: 'assets/checkout_gas_booking.png', caption: 'Stacks & Sections Logic: Seamlessly layering Category-based offers (Gas booking) with Instrument-based nudges (EMI cashback) to maximize value without breaking unit economics.' }
                    ]
                },
                {
                    type: 'technical',
                    title: 'Technical Scale: The Catalogue Service',
                    text: 'To power this experience without degrading payment performance, we overhauled the backend with a dedicated "Catalogue Service." This enabled us to scale seamlessly for massive homepage traffic loads while maintaining strict checkout network latency under 180ms.'
                },
                {
                    type: 'success',
                    title: 'Closing the Loop: Post-Transaction Reward Recall',
                    text: 'Redesigned the success screen to prominently display disbursed benefits with direct redirects to rewards, closing the cognitive loop and building long-term brand trust.',
                    image: 'assets/success.png',
                    caption: 'Post-transaction validation to eliminate “I didn’t get my reward” pain points.'
                }
            ],
            metrics: {
                business: [
                    { label: 'Market Share recovery', value: '+200 bps', class: 'success' },
                    { label: 'NPS Lift', value: '4 Points (67 → 71)', class: 'success' }
                ],
                product: [
                    { label: 'Redemption Lift', value: '67%', class: 'support' },
                    { label: 'Scratch Card engagement', value: '+76%', class: 'support' },
                    { label: 'Checkout Conversion', value: '+22%', class: 'support' },
                    { label: 'Instrument Acquisition', value: '66%', class: 'support' }
                ]
            }
        },
        'merchant-growth': {
            category: 'Platform Strategy @ PhonePe',
            title: 'B2B Merchant Growth & PG Integrations',
            hook: 'Transforming onboarding from a 2-day bottleneck into a 30-minute self-serve engine.',
            narrative: [
                {
                    type: 'intro',
                    title: 'The Operational Bottleneck',
                    text: 'As PhonePe Payment Gateway scaled, onboarding new merchants and configuring complex subvention offers became a terminal operational bottleneck. Manual configuration by internal teams led to a 2-day Turnaround Time (TAT), frustrating high-value merchants and brand partners who needed campaign agility.'
                },
                {
                    type: 'strategy',
                    title: 'Strategic Pivot: Managed Service to Platform Play',
                    text: 'Instead of hiring more operations staff, I spearheaded a pivot to a self-serve infrastructure. My mission was to build intuitive, embeddable subvention consoles that empowered merchants to independently configure and launch offers at scale.'
                },
                {
                    type: 'action',
                    title: 'Platform Execution: The Self-Serve Ecosystem',
                    text: 'We launched Instant Discount and EMI subvention capabilities for both online and offline verticals. I extended these embeddable consoles to other tenants across the group, including the Indus AppStore, allowing third-party developers to seamlessly self-onboard and manage their own offer engines without any PhonePe intervention.'
                },
                {
                    type: 'technical',
                    title: 'Efficiency at Scale',
                    text: 'The architecture was designed for high campaign velocity. By automating the validation and go-live workflows, we slashed the TAT from 48 hours to under 30 minutes, significantly accelerating the time-to-market for merchant marketing campaigns.'
                }
            ],
            metrics: {
                business: [
                    { label: 'Merchants Acquired', value: '5000+', class: 'success' },
                    { label: 'TAT Reduction', value: '(-85%)', class: 'success' }
                ],
                product: [
                    { label: 'Go-Live Velocity', value: '30 Mins', class: 'support' },
                    { label: 'Merchant Retention', value: '+23%', class: 'support' },
                    { label: 'Onboarding Drop-off', value: '< 12%', class: 'support' },
                    { label: 'Scaling Efficiency', value: 'Zero-Touch', class: 'support' }
                ]
            }
        },
        'pincode-growth': {
            category: 'Quick Commerce @ Pincode',
            title: 'Dynamic Cart Incentivization',
            hook: 'Boosting AOV by 35% and dropping cart abandonment by 60% with context-aware triggers.',
            narrative: [
                {
                    type: 'intro',
                    title: 'The Basket Size Problem',
                    text: 'Pincode (PhonePe’s ONDC platform) needed to drive higher basket sizes in a high-density, competitive quick commerce market. Generic coupons were failing to move the needle on Average Order Value (AOV) and retention, leading to sub-optimal unit economics.'
                },
                {
                    type: 'strategy',
                    title: 'Strategic Intervention: Real-Time Intent',
                    text: 'Instead of blanket discounts, we shifted to real-time, cart-context interventions. My strategy was to nudge users at the moment of highest intent—the checkout—based on their specific cart composition.'
                },
                {
                    type: 'action',
                    title: 'Precision Execution',
                    text: 'We launched dynamic triggers like "Add ₹50 more for Free Delivery" and contextual fee waivers based on real-time demand. We also gamified the journey with "Steal Deals" unlocked only at specific order thresholds, driving psychological commitment to larger baskets.'
                },
                {
                    type: 'technical',
                    title: 'Context-Aware Logic',
                    text: 'Built the intervention engine to evaluate cart value and proximity to free-delivery thresholds in real-time, delivering personalized nudges with sub-100ms latency to ensure no friction in the checkout flow.'
                }
            ],
            metrics: {
                business: [
                    { label: 'AOV Lift', value: '+35%', class: 'success' },
                    { label: 'Cart Abandonment', value: '-60%', class: 'success' }
                ],
                product: [
                    { label: 'Retention Rate', value: '+20%', class: 'support' },
                    { label: 'Conversion Lift', value: '+14%', class: 'support' },
                    { label: 'Unit Economics', value: 'Stable', class: 'support' },
                    { label: 'Latency', value: '< 100ms', class: 'support' }
                ]
            }
        },
        'monetization-engine': {
            category: 'Monetization @ PhonePe',
            title: '₹100 Cr/Year ML Rewards Engine',
            hook: 'Pivoting from static distribution to a marketplace ecosystem driving 11% YoY revenue growth.',
            narrative: [
                {
                    type: 'intro',
                    title: 'The Relevance Gap',
                    text: 'Managing a massive portfolio of third-party rewards required a shift from manual curation to algorithmic scaling. Static distribution led to low relevance, high churn among brand partners, and stagnant revenue growth across the ₹100 Cr portfolio.'
                },
                {
                    type: 'strategy',
                    title: 'Strategic Pivot: Static to Dynamic Marketplace',
                    text: 'I spearheaded the transformation of the rewards engine from a fixed bucket model to a scalable, dynamic marketplace ecosystem. The objective: maximize reward utilization by matching high-intent users with relevant advertiser content.'
                },
                {
                    type: 'action',
                    title: 'ML-Driven Personalization',
                    text: 'Architected and deployed recommendation models mapping high-intent users to relevant partners. We refined targeting algorithms to improve Return on Ad Spend (ROAS) for 500+ brands, creating a win-win for both users (relevance) and advertisers (ROI).'
                },
                {
                    type: 'technical',
                    title: 'The Marketplace Core',
                    text: 'Rebuilt the rewards engine as a high-frequency marketplace capable of handling billions of reward allocations. Integrated feedback loops that allowed the ML model to learn from user "scratch card" behavior in real-time to adjust future recommendations.'
                }
            ],
            metrics: {
                business: [
                    { label: 'YoY Revenue Growth', value: '+11%', class: 'success' },
                    { label: 'Advertiser ROAS', value: '+15%', class: 'success' }
                ],
                product: [
                    { label: 'User Engagement', value: '+26%', class: 'support' },
                    { label: 'User Churn Rate', value: '-15%', class: 'support' },
                    { label: 'Reward Redemptions', value: '+42%', class: 'support' },
                    { label: 'Partner Brands', value: '500+', class: 'support' }
                ]
            }
        }
    };

    function openModal(caseId) {
        const data = caseStudies[caseId];
        if (!data) return;

        // Immersive Narrative Rendering for featured case study
        if (data.narrative) {
            modalBody.innerHTML = `
                <div class="modal-header">
                    <div class="modal-category">${data.category}</div>
                    <h1 class="modal-title">${data.title}</h1>
                    <p class="case-summary">${data.hook}</p>
                </div>
                
                <div class="modal-grid">
                    <div class="modal-main">
                        ${data.narrative.map(section => {
                            if (section.type === 'intro') {
                                return `
                                    <div class="star-section">
                                        <h2 class="section-title">${section.title}</h2>
                                        <p>${section.text}</p>
                                        <div class="voc-quote-container">
                                            ${section.voc.map(quote => `<div class="voc-quote">${quote}</div>`).join('')}
                                        </div>
                                    </div>
                                `;
                            }
                            if (section.type === 'figma_breakdown') {
                                return `
                                    <div class="star-section">
                                        <h2 class="section-title">${section.title}</h2>
                                        <p>${section.text}</p>
                                        <div class="comparison-grid">
                                            ${section.images.map(img => `
                                                <div class="before-after-card">
                                                    <div class="modal-image-container">
                                                        <a href="${img.url}" target="_blank" title="View High Resolution">
                                                            <img src="${img.url}" alt="Figma Breakdown">
                                                        </a>
                                                    </div>
                                                    <p class="modal-image-caption">${img.caption}</p>
                                                </div>
                                            `).join('')}
                                        </div>
                                    </div>
                                `;
                            }
                            return `
                                <div class="star-section">
                                    <h2 class="section-title">${section.title}</h2>
                                    <p>${section.text}</p>
                                    ${section.image ? `
                                        <div class="before-after-card" style="margin-top:20px;">
                                            <div class="modal-image-container">
                                                <a href="${section.image}" target="_blank" title="View High Resolution">
                                                    <img src="${section.image}" alt="${section.title}">
                                                </a>
                                            </div>
                                            <p class="modal-image-caption">${section.caption}</p>
                                        </div>
                                    ` : ''}
                                </div>
                            `;
                        }).join('')}
                    </div>
                    
                    <div class="modal-sidebar">
                        <h4 style="margin-bottom: 12px;">Primary Impact</h4>
                        ${data.metrics.business.map(m => `
                            <div class="modal-metric-card success">
                                <div class="modal-metric-value">${m.value}</div>
                                <div class="modal-metric-label">${m.label}</div>
                            </div>
                        `).join('')}

                        <h4 style="margin: 24px 0 12px;">Product Performance</h4>
                        ${data.metrics.product.map(m => `
                            <div class="modal-metric-card support">
                                <div class="modal-metric-value">${m.value}</div>
                                <div class="modal-metric-label">${m.label}</div>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        } else {
            // Standard rendering for other case studies
            modalBody.innerHTML = `
                <div class="modal-header">
                    <div class="modal-category">${data.category}</div>
                    <h1 class="modal-title">${data.title}</h1>
                    <p class="modal-hook">${data.hook}</p>
                </div>
                <div class="modal-main">
                    <p>${data.context}</p>
                    <h4>The Challenge</h4>
                    <p>${data.problem}</p>
                    <h4>Strategic Approach</h4>
                    <ul class="modal-list">
                        ${data.strategy.map(pt => `<li>${pt}</li>`).join('')}
                    </ul>
                    <div class="modal-sidebar" style="position: relative; margin-top: 32px; display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px;">
                        ${data.metrics.map(m => `
                            <div class="modal-metric-card ${m.class}">
                                <div class="modal-metric-type">${m.type}</div>
                                <div class="modal-metric-value">${m.value}</div>
                                <div class="modal-metric-label">${m.label}</div>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        }

        caseModal.classList.add('active');
        document.body.style.overflow = 'hidden'; 
        lucide.createIcons();
    }

    function closeModal() {
        caseModal.classList.remove('active');
        document.body.style.overflow = 'auto';
    }

    readMoreBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const caseId = btn.getAttribute('data-case-id');
            openModal(caseId);
        });
    });

    closeBtn.addEventListener('click', closeModal);
    
    // Close on backdrop click
    caseModal.addEventListener('click', (e) => {
        if (e.target === caseModal) closeModal();
    });

    // Close on Escape key
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeModal();
    });

});

