# /apply — End-to-End Job Application Workflow

> **Invoke this skill whenever Ajay gives you a JD + company name.**
> Work through all 5 phases in order. Never skip a phase. 
> This document is self-evolving — update the [Evolution Log](#evolution-log) after every application run.

---

## 0. GLOBAL AGENT RULES
1. **Never ask for permission** in this project to commit anything to git or make changes. Just execute.
2. **PRDs should always be 1-pager** and almost completely filled.
3. **Only consult for case studies** with Ajay. For everything else, take initiative.
4. **Resume MUST be generated as both HTML AND PDF — always.** After writing the HTML, immediately run the headless Chrome command to produce the PDF. Do NOT declare Phase 5 complete until the `.pdf` file exists on disk. Never ask Ajay to print it himself.
5. **Always cite sources for projected/industry metrics.** Every prototype and CEO deck must include a sources section that maps each projected metric to a named industry benchmark (e.g. Zendesk CX Trends 2024, McKinsey Digital Banking Survey 2022, PYMNTS Payment Failure Index 2023). Sources go in a 2-column grid below the phone frame in prototypes, and inline as "(Source: X)" on impact slides in decks. Never present a projected metric without a benchmark source. Research and cite real publications — do not fabricate sources.

---

## WHO IS AJAY — SOURCE OF TRUTH

Always use this as the single source of truth. Never fabricate metrics or experiences not listed here.

**Ajay Avaghade**
- **Location:** Bangalore, India
- **Phone:** +91 9561558439
- **Email:** avaghadeajay009@gmail.com
- **LinkedIn:** linkedin.com/in/ajay-avaghade
- **Portfolio:** https://ajay-avaghade.github.io/Portfolio/
- **Relocation:** Planning active relocation to Dubai, UAE in Q3 2026. Requires visa sponsorship. *(Include this line only for UAE jobs. Omit entirely for India jobs.)*

### Experience

**Product Manager** — PhonePe · Consumer Offers & Growth Platform *(May 2023 – Present)*
- Inherited a fragmented offers stack across 350M+ MAU with low redemption and high checkout drop-off. Redesigned end-to-end offers placement, eligibility surfacing, and redemption UX → **22% checkout conversion lift**
- ₹1000+ Cr/yr marketing budget spent on static, rule-based targeting. Led deployment of Propensity-to-Transact ML models replacing manual cohorts with real-time user-level scoring → **32% marketing burn reduction**, sustained GMV growth
- Pincode (quick commerce): 60%+ cart abandonment, flat AOV. Built dynamic cart incentivization engine with context-aware triggers (cart value × user intent × time signals) → **35% AOV uplift, 60% cart abandonment reduction, 20% improvement in 30-day retention**

**Associate Product Manager** — PhonePe · Third-Party Rewards & Merchant Growth *(Apr 2022 – Apr 2023)*
- ₹100 Cr/yr rewards portfolio run on static non-personalised rules. Rebuilt into ML-driven marketplace ecosystem with audience segmentation and brand self-serve tooling → **500+ brand partners onboarded, 11% YoY revenue growth, 26% user engagement increase**
- High CAC from undifferentiated merchant acquisition. Built zero-to-one self-serve PG integration platform with Instant Discount & EMI subvention capabilities as the core value proposition → **5,000+ B2B merchants acquired, 23% lower CAC, 5Mn+ new users/month**. Separately, redesigned offer operations workflow → **30-min offer ops TAT** (down from 2 days). *(Note: the 5K merchant acquisition was driven by Instant Discount + EMI subvention platform capabilities — NOT by KYC automation. The 30-min TAT applies to offer operations, not merchant onboarding.)*

**Product Manager — Founding Team (Intern)** — Kotak Mahindra Bank *(Apr 2021 – Jun 2021)*
- Founding team for Kotak Cherry (WealthTech app), zero-to-launch. Led UX research, scoped MVP (asset allocation + goal tracking), co-owned GTM → **100K+ downloads in launch window**

### Education
- **MBA** — Indian Institute of Management (IIM) Indore *(2020–2022)*
- **Exchange Program** — NEOMA Business School, France *(2021)*
- **B.Tech, Computer Science & Engineering** — VNIT Nagpur *(2016–2020)*

### Skills
- **Product & Strategy:** 0→1 Launches, Growth Strategy, Monetization, Marketplace Design, ML Product Management, A/B Testing, Pricing, Q-commerce, New Ventures
- **Data & AI Tools:** SQL, Mixpanel, Amplitude, CleverTap, Metabase, Tableau, Figma, Jira, Vibe-coded Prototyping

### Side Projects
- **AI Mock PM Interviewer** — Conversational AI for PM interview practice with structured scoring and real-time coaching. Link: [AI Mock PM Interviewer](https://huggingface.co/spaces/ajay-avaghade/AI-Mock-Product-Manager-Interviewer)
- **Job Apply AI** — Chrome extension auto-filling job applications on LinkedIn, Workday, Lever, Greenhouse with AI-generated outreach. Link: [Job Apply AI](https://chromewebstore.google.com/detail/antigravity-apply/accbghdhdmdigaohfaggdcnjlifdclbc)
- **Antigravity Resume Agent** — Multi-agent AI pipeline (Gemini) that auto-generates >90 ATS-scored resumes from a JD URL; 6 specialised agents in parallel (scraping, synthesis, layout, ATS scoring, outreach). Link: [Antigravity Resume Agent](https://ajay-avaghade.github.io/Antigravity-Resume-Agent/index.html)

### Awards
- ITC Interrobang — National Finalist (Top 15 teams), ITC's premier pan-India strategy competition
- Asian Paints Canvas — 1st place, pan-India brand strategy & product innovation challenge
- Education Minister Award — State-level academic excellence, Government of Maharashtra

### Existing Case Studies (Portfolio)
| Company | Case Study | Prototype | Deck | Portfolio URL |
|---------|-----------|-----------|------|---------------|
| Talabat | Steal Deals | steal-deals-prototype.html | assets/steal-deals-deck.pptx | #real-world/talabat |
| Talabat | AI Beauty Match | ai-beauty-match-prototype.html | assets/ai-beauty-match-deck.pptx | #real-world/talabat |
| Talabat | Talabat Reels | talabat-reels-prototype.html | assets/talabat-reels-deck.pptx | #real-world/talabat |
| Virtusa | MerchantOS: Milestone Rental Waiver | virtusa-milestone-prototype.html | assets/virtusa-milestone-deck.pptx | #real-world/virtusa |
| Snapmint | SnapPay: Credit on UPI | snapmint-snappay-prototype.html | assets/snapmint-snappay-deck.pptx | #real-world/snapmint |
| Amazon | SmartOffer Engine | amazon-smartoffer-prototype.html | assets/amazon-smartoffer-deck.pptx | #real-world/amazon |
| Amazon | Amazon Pay Milestone | amazon-milestone-prototype.html | assets/amazon-milestone-deck.pptx | #real-world/amazon |
| Ralph Lauren | RL ClientIQ | rl-clientiq-prototype.html | assets/rl-clientiq-deck.pptx | #real-world/ralphlauren |
| Ralph Lauren | RL CartBridge | rl-cartbridge-prototype.html | assets/rl-cartbridge-deck.pptx | #real-world/ralphlauren |
| Agoda | Agoda Spark | agoda-spark-prototype.html | assets/agoda-spark-deck.pptx | #real-world/agoda |
| Agoda | Agoda TripAgent | agoda-tripagent-prototype.html | assets/agoda-tripagent-deck.pptx | #real-world/agoda |
| Agoda | Agoda TrustScore | agoda-trustscore-prototype.html | assets/agoda-trustscore-deck.pptx | #real-world/agoda |
| MultiBank | MultiBank TradeRise | multibank-traderise-prototype.html | assets/multibank-traderise-deck.pptx | #real-world/multibank |
| MultiBank | MultiBank TradeCircle | multibank-tradecircle-prototype.html | assets/multibank-tradecircle-deck.pptx | #real-world/multibank |
| ADCB | ADCB CorporateLens | adcb-corporatelens-prototype.html | assets/adcb-corporatelens-deck.pptx | #real-world/adcb |
| ADCB | ADCB CardLaunch | adcb-cardlaunch-prototype.html | assets/adcb-cardlaunch-deck.pptx | #real-world/adcb |
| PayPay | PayPay SmartBonus + MissionBoard | paypay-smartbonus-prototype.html | assets/paypay-smartbonus-deck.pptx | #real-world/paypay |
| PayPay | PayPay NearPay | paypay-nearpay-prototype.html | assets/paypay-nearpay-deck.pptx | #real-world/paypay |
| J.P. Morgan | J.P. Morgan PayIQ | jpmorgan-payiq-prototype.html | assets/jpmorgan-payiq-deck.pptx | #real-world/jpmorgan |
| J.P. Morgan | J.P. Morgan ClearPath | jpmorgan-clearpath-prototype.html | assets/jpmorgan-clearpath-deck.pptx | #real-world/jpmorgan |
| Adyen | Adyen SmartCheckout | adyen-smartcheckout-prototype.html | assets/adyen-smartcheckout-deck.pptx | #real-world/adyen |
| Adyen | Adyen RegSense | adyen-regsense-prototype.html | assets/adyen-regsense-deck.pptx | #real-world/adyen |
| Navi | Navi GrowthOS | navi-growthos-prototype.html | assets/navi-growthos-deck.pptx | #real-world/navi |
| Navi | Navi CategoryLeap | navi-categoryleap-prototype.html | assets/navi-categoryleap-deck.pptx | #real-world/navi |
| Revolut | Revolut GrowthIQ | revolut-growthiq-prototype.html | assets/revolut-growthiq-deck.pptx | #real-world/revolut |
| Revolut | Revolut PlanPass | revolut-planpass-prototype.html | assets/revolut-planpass-deck.pptx | #real-world/revolut |
| inDrive | inDrive BidBand | indrive-bidband-prototype.html | assets/indrive-bidband-deck.pptx | #real-world/indrive |
| inDrive | inDrive Pass | indrive-pass-prototype.html | assets/indrive-pass-deck.pptx | #real-world/indrive |
| Nykaa | Nykaa QuickStore | nykaa-quickstore-prototype.html | assets/nykaa-quickstore-deck.pptx | #real-world/nykaa |
| Nykaa | Nykaa AdAgent | nykaa-adagent-prototype.html | assets/nykaa-adagent-deck.pptx | #real-world/nykaa |
| Nykaa | Nykaa SellerIQ | nykaa-selleriq-prototype.html | assets/nykaa-selleriq-deck.pptx | #real-world/nykaa |
| Tesco | WalletBridge | tesco-walletbridge-prototype.html | assets/tesco-walletbridge-deck.pptx | #real-world/tesco |
| Tesco | AuthSense | tesco-authsense-prototype.html | assets/tesco-authsense-deck.pptx | #real-world/tesco |
| HealthifyMe | UpgradeIQ | healthifyme-upgradeiq-prototype.html | assets/healthifyme-upgradeiq-deck.pptx | #real-world/healthifyme |
| HealthifyMe | RetentionOS | healthifyme-retentionos-prototype.html | assets/healthifyme-retentionos-deck.pptx | #real-world/healthifyme |
| HealthifyMe | MealMatch | healthifyme-mealmatch-prototype.html | assets/healthifyme-mealmatch-deck.pptx | #real-world/healthifyme |
| Easygenerator | EasyFlow | easygenerator-easyflow-prototype.html | assets/easygenerator-easyflow-deck.pptx | #real-world/easygenerator |
| Easygenerator | EasyExpert | easygenerator-easyexpert-prototype.html | assets/easygenerator-easyexpert-deck.pptx | #real-world/easygenerator |
| Easygenerator | EasyTrust | easygenerator-easytrust-prototype.html | assets/easygenerator-easytrust-deck.pptx | #real-world/easygenerator |
| Rebel Foods | RebelOS | rebelfoods-rebelos-prototype.html | assets/rebelfoods-rebelos-deck.pptx | #real-world/rebelfoods |
| Brain Co. | BrainTrust | brainco-braintrust-prototype.html | assets/brainco-deck.pptx | #real-world/brainco |
| Brain Co. | BrainSync | brainco-brainsync-prototype.html | assets/brainco-sync-deck.pptx | #real-world/brainco |
| Brain Co. | BrainOps | brainco-brainops-prototype.html | assets/brainco-ops-deck.pptx | #real-world/brainco |
| FINN | FinnFlow | finn-finnflow-prototype.html | assets/finn-flow-deck.pptx | #real-world/finn |
| FINN | FinnGuide | finn-finnguide-prototype.html | assets/finn-guide-deck.pptx | #real-world/finn |
| FINN | FinnEarn | finn-finnearn-prototype.html | assets/finn-earn-deck.pptx | #real-world/finn |
| Transguard | CashPulse | transguard-cashpulse-prototype.html | assets/transguard-cashpulse-deck.pptx | #real-world/transguard |
| Transguard | ClientPulse | transguard-clientpulse-prototype.html | assets/transguard-clientpulse-deck.pptx | #real-world/transguard |
| Meesho | ReturnSense | meesho-returnsense-prototype.html | assets/meesho-returnsense-deck.pptx | #real-world/meesho |
| Meesho | Vaani-Cart | meesho-vaanicart-prototype.html | assets/meesho-vaanicart-deck.pptx | #real-world/meesho |
| Meesho | TrustBuy | meesho-trustbuy-prototype.html | assets/meesho-trustbuy-deck.pptx | #real-world/meesho |
| Almosafer | SmartLock | almosafer-smartlock-prototype.html | assets/almosafer-smartlock-deck.pptx | #real-world/almosafer |
| Almosafer | RouteIQ | almosafer-routeiq-prototype.html | assets/almosafer-routeiq-deck.pptx | #real-world/almosafer |
| Almosafer | Milestone Explorer | almosafer-milestone-prototype.html | assets/almosafer-milestone-deck.pptx | #real-world/almosafer |
| Qlub | SmartSplit | qlub-smartsplit-prototype.html | assets/qlub-smartsplit-deck.pptx | #real-world/qlub |
| Qlub | Upsell Engine | qlub-upsell-prototype.html | assets/qlub-upsell-deck.pptx | #real-world/qlub |
| Qlub | AI-Ops Hub | qlub-aiops-prototype.html | assets/qlub-aiops-deck.pptx | #real-world/qlub |
| STYLI | Syndic8 | styli-syndic8-prototype.html | assets/styli-syndic8-deck.pptx | #real-world/styli |
| STYLI | SyncOps | styli-syncops-prototype.html | assets/styli-syncops-deck.pptx | #real-world/styli |
| STYLI | ReconHub | styli-reconhub-prototype.html | assets/styli-reconhub-deck.pptx | #real-world/styli |
| eyewa | eyewa ProLoop | eyewa-proloop-prototype.html | assets/eyewa-proloop-deck.pptx | #real-world/eyewa |
| eyewa | eyewa RepeatIQ | eyewa-repeatiq-prototype.html | assets/eyewa-repeatiq-deck.pptx | #real-world/eyewa |
| Nomupay | MerchantIQ + SupportIQ | nomupay-merchantiq-prototype.html | assets/nomupay-merchantiq-deck.pptx | #real-world/nomupay |
| Nomupay | BatchPayIQ | nomupay-batchpayiq-prototype.html | assets/nomupay-batchpayiq-deck.pptx | #real-world/nomupay |

---

## PORTFOLIO TECHNICAL REFERENCE

**Repo:** https://github.com/ajay-avaghade/Portfolio.git
**Local path:** `/Users/ajayavaghade/Portfolio/`
**Key files:**
- `index.html` — Main portfolio page
- `styles.css` — All CSS (glassmorphism design system)
- `script.js` — Scroll animations, typewriter, promo code
- `assets/` — PPTX decks, build scripts
- `.claude/commands/apply.md` — This skill file

**Design system tokens** (from styles.css):
- Font: Inter (body), Outfit (headings) — Google Fonts
- Accent: `#7c3aed` (purple), Accent glow: `rgba(124,58,237,0.3)`
- Background: `#0f0f1a`, Surface: `rgba(255,255,255,0.03)`
- Text primary: `#f8f8ff`, Muted: `#8b92a5`
- Border: `rgba(255,255,255,0.08)`
- Glass panel: `background rgba(255,255,255,0.04)`, `border 1px solid rgba(255,255,255,0.08)`, `backdrop-filter blur(12px)`

**Real World section pattern (index.html):**
```html
<div class="rw-company-block animate-on-scroll" id="rw-[company-slug]" style="margin-top: 64px;">
  <div class="rw-company-header">
    <div class="rw-company-logo-wrap" style="background: [company-gradient]; border: 2px solid [accent]33;">
      <span class="rw-company-logo-text" style="color: [accent]; font-weight: 800;">[AppName]</span>
    </div>
    <div>
      <h3 class="rw-company-name">[Company Full Name]</h3>
      <p class="rw-company-desc">[1-2 sentence company context + what problem you solved]</p>
    </div>
  </div>
  <div class="rw-case-grid">
    <!-- One .rw-case-card per case study — see existing examples -->
    <div class="rw-case-card glass-panel animate-on-scroll" style="border-top: 3px solid [accent];">
      <div class="rw-problem-tag" style="background: [dark]; color: [accent]; border: 1px solid [accent]44;">[Domain Tag]</div>
      <h3 class="rw-case-title">[Case Study Title]</h3>
      <p class="rw-case-problem">[Problem statement — 2-3 sentences, specific, data-led]</p>
      <div class="rw-metrics-row">
        <div class="rw-metric"><span class="rw-val" style="color:[accent];">[Metric]</span><span class="rw-lbl">[Label]</span></div>
        <!-- 3 metrics total -->
      </div>
      <div class="rw-case-actions">
        <a href="[prototype].html" target="_blank" class="btn btn-primary rw-btn" style="background:[accent]; box-shadow:0 4px 14px [accent-glow];">
          <i data-lucide="smartphone"></i> Live Prototype
        </a>
        <a href="assets/[deck].pptx" download class="btn btn-outline rw-btn" style="border-color:[accent]; color:[accent];">
          <i data-lucide="download"></i> CEO Pitch Deck
        </a>
      </div>
    </div>
  </div>
</div>
```

**Git workflow:**
```bash
git add [specific files only — never git add -A blindly]
git commit -m "concise message"
git push origin main
```

---

## PHASE 1 — JD ANALYSIS & CASE STUDY BRAINSTORM

### Step 1.1 — Parse the JD

Read the full JD and extract:

1. **Role title & seniority** — Exact title, IC vs. manager, level signals
2. **Company priorities (top 3)** — What is the company optimizing for right now? Read between the lines — job descriptions reveal strategy
3. **Must-have experiences** — List every required/preferred qualification the JD calls out
4. **Keywords for ATS** — Every domain term, tool, methodology mentioned (will be used in resume)
5. **Gaps vs. Ajay's background** — Honest assessment of where he's underqualified. Flag these — he may have unlisted experience, or we address the gap in the cover letter
6. **Location** — UAE or India? (Determines relocation note in resume)

Present this as a clean structured summary. Ask Ajay: *"Does this analysis look right? Any context on the company I should know?"*

### Step 1.2 — Brainstorm Case Studies

Generate **3–5 distinct case study concepts** that:
- Address a real, observable problem at this company (do research — news, product teardowns, app store reviews, earnings calls)
- Map directly to the role's priorities (not generic PM concepts)
- Play to Ajay's actual experience (find the parallel to something he's shipped at PhonePe or Kotak)
- Can be demonstrated in a working prototype + told in a CEO deck

For each concept provide:
- **Problem:** What is broken, and why does it matter to the company right now?
- **Proposed solution:** What would Ajay build?
- **Ajay's parallel:** Which PhonePe/Kotak experience is the closest proof point?
- **Prototype-ability:** What 4–5 screens would the prototype show?
- **Wow factor:** What makes this surprising or insightful?

### Step 1.3 — Get Ajay's Input

Present the shortlist and explicitly ask:
1. *"Which direction resonates most with you?"*
2. *"Do you have any personal experience, insight, or data on this problem that would strengthen it?"*
3. *"Do you want to refine the angle or combine ideas?"*

**Do not proceed to Phase 2 until Ajay has confirmed the case study direction.**

### Step 1.4 — Finalise the Brief

Once direction is confirmed, produce a one-page brief:
- Problem statement (with data/evidence)
- Proposed solution (named, with a memorable product name)
- Core mechanics (how it works — the key insight)
- 3 projected impact metrics (tied to what the company cares about)
- Prototype scope (exactly 4 screens, named)
- Deck structure (8 slides, titled)
- Company color palette (primary dark, accent, highlight, white, light gray)

Get Ajay's sign-off: *"Does this brief look right before I start building?"*

---

## PHASE 2 — PROTOTYPE STANDARD

### Design Philosophy

Prototypes must look and feel like **Figma prototypes handed to engineering** — pixel-perfect, realistic, indistinguishable from a real app. The goal is for the hiring manager to believe this shipped.

- **No fixed screen limit.** Use as many screens as needed to tell the full user journey end-to-end. Typical range: 5–8 screens. Don't truncate the story to hit a number.
- **Match the company's real design language** — use their actual colors, typography, spacing, and component patterns. Research their app before building.
- **Single phone, centered** — one 390×844px phone frame centered on the page, with smooth slide/fade transitions between screens. No annotation panels. No multi-phone layouts. The design speaks for itself.
- **Realistic content** — real product names, prices, ratings, timestamps, user names. Never placeholder text.
- **Micro-interactions** — buttons have tap states, bottom sheets animate in, success states have confetti, progress bars fill on entry.

### HTML Prototype Spec

All prototypes live at `/Users/ajayavaghade/Portfolio/[company-slug]-[feature]-prototype.html`

**Page structure:**
```html
<!-- Dark background page, single centered phone -->
<body style="background:#0f0f1a; display:flex; flex-direction:column; align-items:center; min-height:100vh; padding:40px 20px;">
  <div class="proto-header">Company — Feature · Interactive Prototype</div>
  <div class="phone-wrap">
    <!-- phone chrome: 390px wide, 844px tall, border-radius 50px, overflow:hidden -->
    <div class="phone">
      <div class="status-bar"><!-- 9:41, signal, wifi, battery --></div>
      <!-- screens: position:absolute, width:100%, height:100%, transition:transform .35s ease / opacity .3s -->
      <div class="screen active" id="s1">...</div>
      <div class="screen" id="s2">...</div>
      ...
    </div>
  </div>
  <!-- nav dots + prev/next below phone -->
  <div class="nav-row">
    <button class="nav-arr" onclick="go(cur-1)">‹</button>
    <div class="dots" id="dots"></div>
    <button class="nav-arr" onclick="go(cur+1)">›</button>
  </div>
  <!-- screen label -->
  <div class="screen-label" id="slabel">01 — SCREEN NAME</div>
</body>
```

**Screen transitions (slide left/right):**
```js
let cur = 1, total = N;
const LABELS = { 1:'01 — NAME', 2:'02 — NAME', ... };
function go(n) {
  if(n<1||n>total) return;
  const prev = document.getElementById('s'+cur);
  const next = document.getElementById('s'+n);
  const dir = n > cur ? 1 : -1;
  prev.style.transform = `translateX(${-dir*100}%)`;
  prev.style.opacity = '0';
  setTimeout(()=>{ prev.classList.remove('active'); prev.style.transform=''; prev.style.opacity=''; },350);
  next.classList.add('active');
  next.style.transform = `translateX(${dir*100}%)`;
  next.style.opacity = '0';
  requestAnimationFrame(()=>{ next.style.transition='transform .35s ease, opacity .3s'; next.style.transform=''; next.style.opacity='1'; });
  setTimeout(()=>{ next.style.transition=''; },380);
  document.querySelectorAll('.dot').forEach((d,i)=>d.classList.toggle('on',i===n-1));
  document.getElementById('slabel').textContent = LABELS[n];
  cur = n;
  if(n === SUCCESS_SCREEN) startConfetti();
}
// Init dots
const dotsEl = document.getElementById('dots');
for(let i=1;i<=total;i++){
  const d=document.createElement('div'); d.className='dot'+(i===1?' on':''); d.onclick=()=>go(i); dotsEl.appendChild(d);
}
```

**Status bar (always include):**
```html
<div class="status-bar">
  <span class="sb-time">9:41</span>
  <div class="sb-icons">
    <svg><!-- signal --></svg>
    <svg><!-- wifi --></svg>
    <svg><!-- battery --></svg>
  </div>
</div>
```

**Confetti (success screen):**
```js
function startConfetti() {
  const c = document.getElementById('conf');
  if(!c || c.children.length) return;
  const cols = ['#FF9900','#00A8E1','#067D62','#FF6B6B','#FFD700','#4ADE80'];
  for(let i=0;i<40;i++){
    const p=document.createElement('div'); p.className='cp';
    const sz=4+Math.random()*8;
    p.style.cssText=`position:absolute;left:${Math.random()*100}%;top:${-10+Math.random()*30}%;width:${sz}px;height:${sz}px;background:${cols[i%cols.length]};border-radius:${Math.random()>.5?'50%':'3px'};animation:fall ${1.5+Math.random()*2}s ${Math.random()*1.5}s ease-in forwards;`;
    c.appendChild(p);
  }
}
```

**Typography:** Use the company's real font stack. If unknown, default to `-apple-system, 'SF Pro Display', sans-serif` for a native mobile feel. Load via Google Fonts only if the company uses a Google Font.

**Colors:** Use the company's exact brand colors. Research their app/website. Never invent colors.

**QA checklist before declaring prototype done:**
- [ ] All screens render without overlap or cutoff
- [ ] Slide transitions work in both directions
- [ ] Status bar appears on every screen
- [ ] Fonts match the company's real app
- [ ] Confetti fires on the success screen
- [ ] All content is realistic (no Lorem ipsum, no placeholder prices)
- [ ] Scrollable screens use `overflow-y:auto; scrollbar-width:none`
- [ ] Bottom nav (if app has one) is consistent across all screens
- [ ] Touch targets ≥ 44px
- [ ] Colors match the company's real brand

---

## PHASE 3 — CEO DECK STANDARD (PptxGenJS)

### Deck Spec

**Script location:** `/Users/ajayavaghade/Portfolio/assets/build-[company-slug]-deck.js`
**Output:** `/Users/ajayavaghade/Portfolio/assets/[company-slug]-deck.pptx`
**Run:** `node assets/build-[company-slug]-deck.js`

**Dependencies:** `npm install pptxgenjs` (already installed in Portfolio dir)

**Layout:** Always `LAYOUT_16x9` (10" × 7.5")

**Color constants pattern:**
```js
const C = {
  dark:   '[primary-hex-no-hash]',  // e.g. '0A1F0A'
  hero:   '[slightly-lighter]',
  bright: '[accent]',
  orange: '[highlight]',
  white:  'FFFFFF',
  lgray:  '[light-gray]',
};
// Shadow factory — ALWAYS a function, never reuse the same object
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });
```

**8-slide structure (mandatory):**

| # | Slide Name | Background | Key Elements |
|---|-----------|-----------|--------------|
| 1 | Cover | Dark (company primary) | Product name, subtitle, presenter line, company tag, bold visual motif (diagonal lines / burst), big stat bottom-right |
| 2 | The Problem | Dark | Slide eyebrow "THE PROBLEM", H1 with core stat in title, 3 stat columns (icon + big number + label + sub-text), insight box at bottom |
| 3 | The Insight | Light | "THE INSIGHT", H1 framing the reframe, 2-column comparison (❌ current vs ✅ proposed), VS divider oval, callout box |
| 4 | The Mechanic | Dark | "THE MECHANIC", H1, numbered timeline steps (5 steps) with connecting dashed line, A/B test proof note at bottom |
| 5 | The Product | Light | "THE PRODUCT", H1, 4 card grid (one per screen), screen number label + title + description + ASCII mockup block |
| 6 | Impact & ROI | Dark | "IMPACT & ROI", 2×4 metric grid (left: Merchant/User impact; right: Company/Acquirer ROI), insight box at bottom |
| 7 | Proof of Work | Light/split | "PROOF OF WORK", 2 columns (left dark: what I built at PhonePe; right light: how it maps to this JD), closing quote box |
| 8 | Rollout Plan | Dark | "ROLLOUT PLAN", 3 phase cards, "What I need" ask box at bottom, contact footer |

**Typography:**
- Slide titles (eyebrow): 9–10pt, ALL CAPS, 0.08em letter-spacing, muted
- H1 (main heading): 26–30pt, bold, high contrast
- Body / metric labels: 13–14pt
- Big stat numbers: 36–44pt, bold, accent color
- Footer / captions: 9–10pt

**Design rules:**
- NEVER use underline accent lines under titles (AI cliché)
- Use background color blocks or whitespace to separate sections instead
- Every slide must have at least one visual element — shape, colored block, or icon character
- Dark slides: white text on dark background; Light slides: dark text on white/light background
- Alternate dark/light ("sandwich" structure) — Cover dark, Problem dark, Insight light, Mechanic dark, Product light, ROI dark, Proof light, Rollout dark

**QA (mandatory before declaring done):**
```bash
python3 -c "
from pptx import Presentation
prs = Presentation('assets/[company]-deck.pptx')
for i, slide in enumerate(prs.slides, 1):
    print(f'--- SLIDE {i} ---')
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = para.text.strip()
                if t: print(t)
"
```
Check for: missing slides, leftover placeholder text (XXX, TODO, lorem), wrong slide count (must be 8).

---

## PHASE 4 — PORTFOLIO INTEGRATION

After prototype and deck are confirmed:

1. **Add company block to index.html** following the template in [Portfolio Technical Reference](#portfolio-technical-reference)
   - `id="rw-[company-slug]"` on the company block — enables deep-link URL
   - `border-top: 3px solid [accent]` on case card
   - Buttons: `class="btn btn-primary rw-btn"` and `class="btn btn-outline rw-btn"` — ALWAYS use this exact class structure, never invent new classes
   - 3 impact metrics in `rw-metrics-row`

2. **Test the deep-link URL** (hash router already handles it):
   `https://ajay-avaghade.github.io/Portfolio/#real-world/[company-slug]`

3. **Git commit:**
```bash
git add index.html [prototype].html assets/[deck].pptx assets/build-[company]-deck.js
git commit -m "Add [Company] [Feature] case study — prototype + CEO deck"
git push origin main
```

4. **Update the Existing Case Studies table** in this skill file (Section: WHO IS AJAY → Existing Case Studies)

---

## PHASE 5 — RESUME GENERATION

### Resume Spec

**Outputs (both required — Phase 5 is not complete until both files exist):**
- `/Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].html` — source file
- `/Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].pdf` — generated automatically via headless Chrome immediately after the HTML is written

**PDF generation command (run this automatically, no prompting):**
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --headless --no-pdf-header-footer --print-to-pdf="/Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].pdf" "file:///Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].html"
```
Then open it: `open "/Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].pdf"`

**One-pager rule — STRICT:** Everything must fit on a single A4/Letter page. The page must also be **completely filled** — content should reach close to the bottom margin, not end halfway down. After writing, mentally estimate page fill: if content ends above ~85% of the page, add substance (expand a bullet, add a relevant case study to Projects). If overflowing: reduce `@page` margins first (.20in .25in), then trim bullet text — do NOT shrink font below 8.2pt.

**CASE STUDY SEPARATION RULE — CRITICAL:**
- **Experience bullets** must describe what was built and the results — they must NEVER reference or name the case study being built for this application. The hiring manager will see Experience and Case Studies as two separate proof points; conflating them kills that effect.
- **Professional Summary** must end with a brief case study signal — just the name + a 5-word description max. Example: *"Built two Revolut-specific products — LoyaltyLoop and UpgradeOS — as applied case studies."* No metrics, no mechanics in the summary.
- **Case Studies section** = name + ONE sentence (what problem it solves) + [prototype] [deck] links. Nothing more.

**Title adaptation:** Rewrite the headline title to mirror the JD's language exactly.
- If JD says "Senior Product Manager, Growth" → title is "Product Manager — Growth & Monetization"
- If JD says "Head of Product, Merchant" → title is "Product Manager — Merchant Products & Growth"

**STAR framework for every bullet:**
Every experience bullet must follow: **Situation → Action → Result**
- Situation (15 words max): The problem or context
- Action (20 words max): What Ajay specifically owned/built
- Result (10 words max, lead with the metric): The outcome

Format: `[Situation context]; [action taken] — [metric result].`

**ATS optimisation:**
1. Extract every keyword, tool, methodology from the JD (Phase 1 output)
2. Weave ALL of them naturally into: Professional Summary (must contain top 5–7 keywords), Skills section (verbatim match), bullet points (use JD's exact phrasing where authentic)
3. Never keyword-stuff — every keyword must appear in a sentence that makes sense
4. Target: ATS score > 90. Check by mentally scanning: does every required qualification have a corresponding resume line?

**Mandatory sections and order:**
1. Name (large, centered, bold)
2. Role title (centered, italic, adapts to JD)
3. Contact line (city, phone, email, LinkedIn, Portfolio) — centered
4. Relocation line (UAE jobs only): *"Relocating to Dubai, UAE · Q2 2026 · Visa Sponsorship Required"*
5. PROFESSIONAL SUMMARY — 3–4 sentences. Opens with category label + years. Mentions the company applied to's domain. Ends with the case study signal.
6. EXPERIENCE — Reverse chronological. STAR bullets. 2–3 bullets per role.
7. EDUCATION — IIM, NEOMA, VNIT
8. SKILLS — Two rows only: Product & Strategy, Data & AI Tools
9. CASE STUDIES — [Company Applied To] (this is the case study just built)
   Format per case study: `• [Title] — [1-sentence what was solved]. [prototype] [deck]`
10. PROJECTS & SIDE BUILDS — Do NOT include all 3 side projects by default. Depending on the remaining vertical space on the page (the resume must be exactly one page and completely filled to >85% of the page), choose how many side projects to include. The sequence of picking projects MUST be: (1) AI Mock PM Interviewer, (2) Job Apply AI, and (3) Antigravity Resume Agent. Make them hyperlinked using HTML anchor tags pointing to their URLs: AI Mock PM Interviewer (https://huggingface.co/spaces/ajay-avaghade/AI-Mock-Product-Manager-Interviewer), Job Apply AI (https://chromewebstore.google.com/detail/antigravity-apply/accbghdhdmdigaohfaggdcnjlifdclbc), Antigravity Resume Agent (https://ajay-avaghade.github.io/Antigravity-Resume-Agent/index.html). Additionally, include 1-2 most relevant portfolio case studies (excluding the current company's).
    Format each project as a bullet with an anchor tag:
    - `• <a href="https://huggingface.co/spaces/ajay-avaghade/AI-Mock-Product-Manager-Interviewer" target="_blank">AI Mock PM Interviewer</a> — Conversational AI for structured PM interview practice with real-time coaching feedback.`
    - `• <a href="https://chromewebstore.google.com/detail/antigravity-apply/accbghdhdmdigaohfaggdcnjlifdclbc" target="_blank">Job Apply AI</a> — Chrome extension auto-filling job applications on LinkedIn, Workday, Lever, Greenhouse; AI-generated outreach emails.`
    - `• <a href="https://ajay-avaghade.github.io/Antigravity-Resume-Agent/index.html" target="_blank">Antigravity Resume Agent</a> — Gemini-powered multi-agent pipeline (6 parallel agents) auto-generating >90 ATS-scored resumes from a JD URL.`
11. AWARDS & RECOGNITION — Always include all three:
    - ITC Interrobang — National Finalist, Top 15 teams
    - Asian Paints Canvas — 1st place
    - Education Minister Award — State-level academic excellence, District Topper (10th & 12th)
    *(Do NOT append "Government of Maharashtra" — it causes the awards line to overflow onto page 2.)*

**HTML Resume CSS baseline (proven working values — do not inflate these):**

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Ajay Avaghade — Resume</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&display=swap');
  * { margin:0; padding:0; box-sizing:border-box; }
  body {
    font-family:'Inter',Arial,sans-serif;
    font-size:8.4pt;
    color:#1a1a1a;
    background:#fff;
    padding:14px 26px;
    max-width:800px;
    margin:0 auto;
    line-height:1.32;          /* PROVEN: 1.35 overflows — do not increase */
  }
  .name       { font-size:19pt; font-weight:700; text-align:center; letter-spacing:.04em; }
  .role-title { font-size:9pt; font-style:italic; text-align:center; color:#444; margin:1px 0 2px; }
  .contact    { font-size:7.6pt; text-align:center; color:#444; margin-bottom:4px; }
  .contact a  { color:#1a56db; text-decoration:none; }
  .relocation { font-size:7.6pt; font-style:italic; text-align:center; color:#555; margin-bottom:3px; }
  hr { border:none; border-top:1.5px solid #1a56db; margin:3px 0 2px; }
  h2 { font-size:7.8pt; font-weight:700; color:#1a56db; text-transform:uppercase; letter-spacing:.07em; margin:4px 0 1px; }
  .job-row   { display:flex; justify-content:space-between; align-items:baseline; }
  .job-title { font-size:8.8pt; font-weight:700; }
  .job-date  { font-size:7.6pt; color:#555; white-space:nowrap; }
  .company   { font-size:7.6pt; color:#555; font-style:italic; margin-bottom:1px; }
  ul { padding-left:12px; margin:1px 0 2px; }
  li {
    font-size:8.2pt;
    line-height:1.36;          /* PROVEN: 1.38 overflows — do not increase */
    margin-bottom:1px;
    text-align:justify;        /* fills middle lines edge-to-edge */
    text-align-last:left;      /* last line stays natural — kills orphan-word stretching */
  }
  .summary {
    font-size:8.2pt; line-height:1.38;
    text-align:justify; text-align-last:left;
  }
  .skills-row {
    font-size:8.1pt; line-height:1.38; margin-bottom:1px;
    text-align:justify; text-align-last:left;
  }
  .skills-row strong { font-weight:600; color:#0f0f0f; }
  .edu-row  { display:flex; justify-content:space-between; align-items:baseline; }
  .edu-name { font-size:8.6pt; font-weight:600; }
  .edu-date { font-size:7.6pt; color:#555; }
  .edu-sub  { font-size:7.6pt; color:#555; font-style:italic; margin-bottom:2px; }
  a { color:#1a56db; font-weight:500; text-decoration:none; }
  .awards-line {
    font-size:8.1pt; line-height:1.38;
    text-align:justify; text-align-last:left;
  }
  @media print {
    body { padding:0; }
    @page { margin:.18in .26in; size:letter; }  /* PROVEN tight — do not increase */
  }
</style>
</head>
<body>
  <!-- Fill all sections per spec above -->
</body>
</html>
```

**Bullet writing rules — CRITICAL for avoiding orphan lines:**

The biggest layout killer is 1–3 words dangling alone on a bullet's last line. `text-align-last:left` stops them from stretching, but they still waste a full row of whitespace.

Rules to prevent this:
1. **Bold metric phrases must appear MID-sentence, never as the trailing words.** End every bullet with a natural closing phrase after the metric.
   - ❌ `...cut marketing burn by **32%**` ← "32%" alone on last line
   - ✅ `...cut **marketing burn by 32%** while sustaining GMV across all channels.`
2. **Last line of every bullet must have ≥ 6 words.** Count mentally before finalising.
3. **If a bullet ends with a short fragment, add context** — "across the active user base", "versus the prior manual process", "at the transaction level" are valid closes.
4. **Awards line**: end the Education Minister Award at "District Topper (10th & 12th)" — do NOT append "Government of Maharashtra" as it causes page overflow. The ITC entry can be padded with "out of 5,000+ entries" if line-fill is needed.

---

### Resume QA — Run as Senior Tester Before Handing Over

After writing the resume, **always open it in Chrome and check these as a senior QA tester — do not skip:**

**Page check:**
- [ ] Exactly 1 page in Chrome print preview (⌘P) — no overflow, no excessive blank space at bottom
- [ ] Page is **completely filled** — content reaches within ~10% of the bottom margin. If it ends halfway, add substance before shipping.
- [ ] If overflowing by a few lines: reduce `@page` margins first (.20in .25in), then trim bullet text — do NOT shrink font below 8.2pt

**Case study separation check:**
- [ ] No experience bullet names or describes the case study built for this application
- [ ] Professional summary ends with case study name + ≤5-word description only — no metrics, no mechanics
- [ ] Case studies section = name + 1 sentence + links. Not a paragraph.

**Orphan line check (scan every bullet):**
- [ ] No bullet's last line has fewer than 6 words
- [ ] No bold metric phrase is the very last thing on a bullet (add a closing clause after it)
- [ ] Awards paragraph doesn't end with a single standalone word

**Content check:**
- [ ] Role title mirrors JD language exactly
- [ ] Top 7 JD keywords appear in summary + skills
- [ ] Every bullet is STAR-structured (situation; action — result)
- [ ] Case study section links to prototype + deck
- [ ] UAE relocation line: present for UAE jobs, absent for India jobs
- [ ] No metrics invented — everything in "WHO IS AJAY" section
- [ ] All links work (Portfolio, LinkedIn, prototype, deck)

**After completing the QA checklist above, immediately run the PDF generation command — this is mandatory, not optional:**

```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --headless --no-pdf-header-footer --print-to-pdf="/Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].pdf" "file:///Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].html"
```

Then open the PDF so Ajay can see it:
```bash
open "/Users/ajayavaghade/Job Apply AI/Ajay_Avaghade_Resume_[Company].pdf"
```

**Phase 5 is NOT complete until:**
1. The `.html` file exists at the correct path
2. The `.pdf` file exists at the correct path (same name, `.pdf` extension)
3. The PDF has been opened for Ajay to review

**Quick-edit tip (always share this):**
> **To edit the resume:** Open the `.html` file in Chrome → press ⌘⌥I → Console → type `document.body.contentEditable='true'` → Enter. Click any text and edit inline. Then re-run the Chrome headless command to regenerate the PDF.

---

## PHASE 6 — OUTREACH (LinkedIn + Email)

**File Format**: ALWAYS write outreach messages into a plain text file (`/Users/ajayavaghade/Job Apply AI/[Company]_Outreach.txt`). Do NOT use markdown syntax (like `**bold**` or `[link](url)`) as it requires manual sorting.

### A. LinkedIn Connection Note (300 chars max)

When sending a connection request, the note must:
- Mention the case study by name
- State the specific role
- Have a hook in the first 8 words
- Include a clear, low-friction ask

**Template:**
```
Hi [Name], I solved [Company's specific problem] as a case study — built a working prototype + CEO deck. Applying for [Role]. Would love your take or a referral if it fits. Portfolio: [link]
```
Trim to ≤300 chars. Cut words, not ideas.

### B. Email (for hiring managers or referrals via warm intros)

**Subject line options (choose based on tone of company):**
- `Re: [Role] — I built a [Company] case study before applying`
- `[Role] application — prototype + CEO deck attached`
- `[Name] → [Ajay] — quick note on the [Role] opening`

**Email body structure:**
```
Hi [Name],

[Hook — 1 sentence. The case study, not the resume.]
I spent time studying [Company]'s [specific problem] and built a working prototype + CEO pitch deck to show how I'd solve it — [case study name].

[Why it's relevant — 2 sentences. Map 2 PhonePe experiences to 2 JD priorities.]
At PhonePe, I [parallel 1 — verb + outcome]. Before that, [parallel 2 — verb + outcome]. Both map directly to what [Company]'s [Role] is optimizing for.

[The ask — 1 sentence. Be specific and easy to say yes to.]
Would you be open to a 20-minute call, or if this fits your team's needs, a referral would mean a lot?

Prototype: [URL]
Portfolio: https://ajay-avaghade.github.io/Portfolio/#real-world/[company-slug]
Resume: attached

Best,
Ajay Avaghade
+91 9561558439
```

**Email QA:**
- [ ] Subject line has the company name and role (for easy search)
- [ ] Hook is the case study, not "I'm excited to apply"
- [ ] Two specific PhonePe parallels — not vague ("I have experience in growth")
- [ ] Ask is one clear action (call OR referral, not both)
- [ ] Case study URL points to the new portfolio deep link

---

## INTERACTION RULES

These rules govern how this skill runs every time:

1. **Always confirm the case study direction before building anything.** Prototypes and decks take effort — never start without Ajay's sign-off on Phase 1.4 brief.

2. **Ask targeted questions, not open-ended ones.** Instead of "What do you think?", ask "Between option A (flash deals mechanic) and option B (loyalty tier), which maps better to your PhonePe experience?"

3. **Flag gaps honestly.** If the JD asks for fintech regulatory experience and Ajay doesn't have it, say so. Offer a framing that closes the gap (adjacent experience, side project, or the case study itself).

4. **One deliverable at a time.** Don't dump prototype + deck + resume + email in one message. Deliver Phase by Phase and wait for confirmation.

5. **Never fabricate metrics.** If Ajay hasn't given you a number, use a range ("15–20% improvement") or note it as projected ("estimated 25% uplift based on PhonePe baseline").

6. **Self-update this file.** After each application run, add a row to the Evolution Log below and update the Existing Case Studies table.

7. **Provide Links Promptly.** After every deliverable (prototype, deck, resume, etc.), provide Ajay with the direct local file links (using `file:///`) so he can click and open them. Specifically, after integrating the case study into the portfolio, immediately give him the deep-link to that section (e.g., `https://ajay-avaghade.github.io/Portfolio/#real-world/[company-slug]`).

8. **Always Push to Portfolio.** Whenever you build a case study (prototype + deck), you MUST automatically proceed to Phase 4 (Portfolio Integration) and push it to the portfolio site via git, unless Ajay explicitly specifies not to add it.

---

## EVOLUTION LOG

*This section updates after every application run. It captures what worked, what didn't, and improvements to the workflow.*

| Date | Company | Role | Case Study | Outcome | Learnings |
|------|---------|------|-----------|---------|-----------|
| 2025-05 | Talabat | New Ventures PM | Steal Deals, AI Beauty Match, Talabat Reels | Portfolio shipped | Established the multi-screen phone mockup format + PptxGenJS 8-slide structure |
| 2025-05 | Virtusa | Merchant Acquiring PM | MerchantOS: Milestone Rental Waiver | Portfolio shipped | Discovered `.ns` class collision bug in prototypes — sensor dot vs. card text. Added to QA checklist. Also: button class structure must be `btn btn-primary rw-btn`, never invent new classes. Hash router added to portfolio for deep links. |
| 2025-05 | Snapmint | UPI PM | SnapPay: Credit on UPI | Portfolio shipped | Combined two case study concepts (A+D) into one end-to-end product story — UPI Credit Line setup + contextual credit at QR intercept. Used `.ndot` for notch sensor to avoid `.ns` collision. India job — no relocation line. Renamed company slug: `rw-snapmint`. |
| 2025-05 | Snapmint | UPI PM | Resume QA learnings | Resume finalised |
| 2025-05 | Amazon | Payments PM | SmartOffer Engine + Amazon Pay Milestone | Portfolio shipped | First dual-case-study run. SmartOffer = propensity ML suppression for IBD; Milestone = 3-tier habit loop bank-partner funded. Both prototype + deck + portfolio card. India job — no relocation note. Portfolio slug: rw-amazon. Apostrophes in JS single-quoted strings cause SyntaxError — always escape or rephrase. | Critical resume rules locked in: (1) `text-align:justify; text-align-last:left` on ALL text elements — kills orphan-word stretching on last lines; (2) bold metric phrases must be MID-sentence, never trailing; (3) every bullet's last line must have ≥6 words; (4) `@page margin:.22in .28in` is the proven tight-but-safe print margin — do not increase; (5) always open in Chrome + provide contentEditable quick-edit tip after every resume delivery. |
| 2026-05 | Ralph Lauren | Retail Tech PM | RL ClientIQ + RL CartBridge | Portfolio shipped | Luxury brand — cream/ivory phone backgrounds, navy headers, gold (#C9A84C) accents, burgundy (#7D2027) problem tags. ClientIQ = Customer 360 on QR scan (LTV, wishlist, sizes, history). CartBridge = omnichannel stock-out recovery (POS alert → inventory map → ship-to-home in 47s). Both prototype + deck + portfolio block (slug: rw-ralphlauren). Resume tailored to retail tech + clienteling + omnichannel keywords. |
| 2026-05 | Agoda | Growth/New Verticals PM (Bangkok) | Agoda Spark + TripAgent + TrustScore | Portfolio shipped | 3 case studies all user-ideated (not AI-proposed). Agoda palette: dark navy #0B1F3A / orange #FF6B35 / gold #FFB800 / green #22C55E / purple #8B5CF6 for cab expansion. Spark = shoppable creator video → trip bundle (35% session-to-booking lift). TripAgent = AI agent for end-to-end trip planning + cab new vertical ($52 cab GMV/trip in purple). TrustScore = persona-weighted hotel scoring (Solo/Couples/Family/Business) + ML fake review detection (78K flags/mo). Bangkok job — NOT UAE — no relocation note in resume. Portfolio slug: rw-agoda (3 case cards). |
| 2026-05 | MultiBank Group | Growth Marketing PM (Dubai, UAE) | TradeRise + TradeCircle | Portfolio shipped | UAE job — relocation note included. Palette: dark navy #0B1A2E / teal #00D4AA / gold #F5A623. TradeRise = 30-day milestone activation (5 behaviour-linked rewards: spread discount, rebate, leverage upgrade — same MerchantOS architecture). TradeCircle = guilt-free viral referral: anonymised % return card, circle-based mechanic, success-linked rebates (guilt is the hidden blocker in trading referrals — the key insight). TradeRise accent teal, TradeCircle accent gold. Resume highlights PLG, referral programs, viral mechanics, gamification, CRM/reactivation, CAC/LTV. Virtusa MerchantOS listed as extra proof in Projects (same architecture as TradeRise). |
| 2026-05 | ADCB | PM Corporate Cards & Transaction Banking (Abu Dhabi, UAE) | CorporateLens + CardLaunch | Portfolio shipped | UAE job — relocation note: "Relocating to Abu Dhabi / Dubai, UAE · Q2 2026 · Visa Sponsorship Required". Palette: #00302A dark / #00453C hero / #00BFA5 teal / #D4A843 gold. CorporateLens = real-time self-serve card control plane (limit change 48 hrs → 8 sec, MCC toggles, scheme compliance). CardLaunch = 4-step wizard: OCR + AI credit scoring + virtual card issuance (3 weeks → 24 min). Key proof-point mapping: 2-day → 30-min PhonePe B2B onboarding maps exactly to CardLaunch. 8-digit hex alpha codes (e.g. #00BFA544) throw warnings in pptxgenjs but deck still generates — cosmetic issue only. Resume title: "Product Manager — Corporate Cards, Transaction Banking & Digital Onboarding". Skills reframed toward corporate card domain, Visa/Mastercard scheme compliance, KYC/KYB, cross-functional delivery with Compliance/Legal/Finance. |
| 2026-05 | J.P. Morgan | CEEMEA Core Cash Payments PM (Dubai, UAE assumed) | PayIQ + ClearPath | Portfolio shipped | UAE job — relocation note included. Palette: #0A1E3D dark / #0F2A54 hero / #2563EB blue / #C9A227 gold / #0D9488 teal. PayIQ = ML multi-variable rail scoring (amount × corridor × urgency × cut-off → SARIE / AANI / SWIFT recommendation) + real-time gpi + ISO 20022 client dashboard. ClearPath = correspondent KYC renewal self-serve portal: OCR auto-fill from prior submission + parallel 4-jurisdiction compliance review (UK / UAE / US / local) → 6–9 months → 38 days. Key gap bridge: institutional wholesale payments vs Ajay B2C background → PhonePe context-aware trigger engine = PayIQ rail scoring; PhonePe B2B onboarding + ADCB CardLaunch = ClearPath parallel KYC. Resume title: "Product Manager — CEEMEA Core Cash Payments, Payment Strategy & Client Solutions". Skills: Core Cash Payments, Cross-Border Rail Selection (SWIFT/SARIE/AANI/RTGS/ISO 20022), P&L Management, B2B KYC/KYB Digitisation, Risk Management, Thought Leadership. Apostrophe fix: "J.P. Morgan's" inside single-quoted JS string caused SyntaxError — rephrased to avoid contraction. |
| 2026-06 | Qlub | Senior Product Manager (Dubai, UAE) | SmartSplit, Upsell Engine, AI-Ops Hub | Portfolio shipped | UAE job — relocation note included. Palette: Electric Violet #7D00D4 / Cod Gray #0C0C0C. 3 separate case studies generated spanning B2C checkout (SmartSplit), digital ordering upsell (Upsell Engine) and B2B reconciliation (AI-Ops Hub). Encountered and fixed pptxgenjs version 3.x ShapeType issue by standardising shape constants (pres.ShapeType.rect) and fill object formatting. |
| 2026-06 | STYLI | Product Manager - Marketplace-as-a-Service (MaaS) (Dubai, UAE) | Syndic8, SyncOps, ReconHub | Portfolio shipped | UAE job — relocation note included. Palette: #0F172A dark / #3B82F6 blue / #F8FAFC bg. Generated 3 case studies mapping PhonePe B2B Merchant Growth directly to STYLI MaaS role requirements (Syndic8 for catalog syndication, SyncOps for inventory buffers, ReconHub for financial reconciliation). Evaluated 2 STYLI JDs and chose the MaaS role due to stronger B2B integration alignment over the Influencer role. |
| 2026-05 | Adyen | EMEA Payments PM (Amsterdam, Netherlands) | SmartCheckout + RegSense | Portfolio shipped | Netherlands job — relocation line: "Actively relocating to the Netherlands · Q2 2026 · Visa Sponsorship Required" (custom — not UAE standard, not omitted). Palette: #0D1F12 dark / #102218 hero / #00BB59 Adyen green (SmartCheckout) / #0D9488 teal (RegSense) / #F97316 orange / #F59E0B amber. SmartCheckout = ML-powered EMEA LPM ranking using 23 session signals (locale, device, cart, time, history); ranks iDEAL, Bancontact, Bizum, Giropay per session — no merchant dev required; Screen 3 is consumer-facing light/white to show checkout UX from end-user perspective. RegSense = PSD3/PSR3 proactive compliance intelligence; monitors EBA/regulator changes, calculates merchant-specific exposure, sends 60-day advance alert with numbered resolution path + SCA exemption optimiser benchmarked vs Adyen top-quartile peers + ops portfolio health view. Key gap bridge: no direct EMEA experience → deep EMEA LPM knowledge embedded in SmartCheckout data (iDEAL, Bancontact, Bizum, Giropay with real conversion benchmarks); PSD3/PSR3 regulatory knowledge demonstrated through RegSense mechanic (SCA, TRA, MIT consent, payee whitelisting) — differentiator vs candidates without regulatory depth. Resume title: "Product Manager — EMEA Payments, Checkout Conversion & Regulatory Intelligence". Skills: EMEA LPMs, ML product management, PSD3/PSR3/SCA exemptions, checkout conversion, cross-functional delivery. Notch sensor: always .ndot (never .ns — class collision bug). Apostrophes in JS single-quoted strings cause SyntaxError — always rephrase. |
| 2026-05 | Navi | Growth PM — MarTech & CLTV (Bangalore, India) | GrowthOS + CategoryLeap | Portfolio shipped | India job — NO relocation note. Palette: #071A12 dark / #0E2A1C hero / #00C853 bright green / #FFB800 gold / #FF6B35 orange. GrowthOS = CDP-powered MarTech intelligence platform: propensity banding (HIGH >0.7 no offer / MED 0.4–0.7 minimum calibrated offer / LOW <0.4 suppressed) per product (UPI / Loans / Insurance), story variant engine (Aspiration / Fear / Social Proof), offer calibration engine saves ₹106 Cr per cycle (₹142 Cr flat → ₹36 Cr precision). CategoryLeap = cross-category milestone engine: UPI + Loans + Insurance pillar milestones contribute to shared Portfolio Score, unlocking Bronze → Silver → Gold → Platinum tiers; 3.7× CLTV uplift (1-cat ₹8.4K → 3-cat ₹61.2K). Key correction mid-Phase-1: Navi is a distribution layer — does NOT own loan books or credit underwriting. All ML must target marketing/engagement layer, not credit decisioning. CategoryLeap phone labels in gold (#FFB800) to differentiate from GrowthOS green (#00C853). Notch sensor: .ndot (never .ns). Both prototypes built before decks — confirmed CLTV data: 1-cat ₹8.4K, 2-cat ₹24.6K, 3-cat ₹61.2K. Resume title: "Product Manager — Growth, MarTech & CLTV Products". Skills: Growth MarTech, CDP Platform Design, ML Propensity Modelling, Offer Calibration, Cross-Category CLTV, Milestone & Loyalty, CRM & Story Variant Engine. |
| 2026-05 | PayPay | Senior PM Growth Platform (Tokyo, Japan) | SmartBonus + NearPay | Portfolio shipped | Japan job — NO relocation note. Palette: #1A0A0A dark / #2D1515 hero / #E60012 red / #FF6B00 orange-gold / #00B894 teal (NearPay). SmartBonus = ML propensity score per user sets personalised bonus rate, MissionBoard layered on top lets user earn higher rate via spending challenges (combined into one flywheel case study). NearPay = GPS + time-of-day + category affinity triggers Discovery Bonus push within 200m of unvisited merchant; 100% merchant-funded bonus = zero net cost to PayPay. Key insight: SmartBonus mechanic is literally the same propensity-to-transact model Ajay shipped at PhonePe — applied to bonus rate assignment instead of offers. Phone screens: light white UI with red headers (matching PayPay app aesthetic — opposite of dark-themed prototypes). Resume title: "Product Manager — Growth Platform, Engagement & Merchant Products". Skills: ML Product Management, Engagement & Rewards Systems, Merchant Ecosystem, Gamification & Mission Mechanics. Management gap framing: cross-functional delivery at 350M MAU as horizontal coordination equivalent. Japanese language gap: acknowledge, frame 50+ country team as signal they value international perspective. |
| 2026-05 | inDrive | Senior PM — Growth & Personalization (Cyprus) | BidBand + inDrive Pass | Portfolio shipped | Not UAE, not India — no relocation note. Palette: #0B0E13 dark / #B5E033 lime (BidBand) / #FF6B35 orange (Pass). BidBand: 4 fare bands (Try Your Luck 20% / Fair Offer 45% / Good Chance 72% / Fast Match 91%) replacing single suggested price — eliminates jarring post-no-match jump. Key insight: exclusive to inDrive (Uber/Bolt have no negotiation layer). Waiting state: driver feed (viewed/countered/declined) + nudge card to move to Band 3. Pass: monthly subscription, % back every ride as credits (Bronze 5% / Silver 8% / Gold 12% / Explorer 15%), tier by ride count + cross-vertical (Rides+Freight+Services = Explorer). Case study from Ajay's personal UX frustration — authentic insight. Resume title: "Senior Product Manager — Growth Loops, Personalization & Lifecycle Orchestration". |
| 2026-06 | Transguard Group | Manager — Products, Cash Services (Dubai, UAE) | CashPulse + ClientPulse | Portfolio shipped | UAE job — relocation note: "Relocating to Dubai, UAE · Q3 2026 · Visa Sponsorship Required". Palette: #0A1A2E dark / #0F2540 hero / #00BCD4 teal (CashPulse) / #F59E0B amber (ClientPulse) / #F5A623 gold. CashPulse = real-time SCDM intelligence platform: deposit confirmations <5 min (vs 24 hr SAP batch), live fill levels per machine, ERP auto-match engine (SAP/Oracle/Dynamics), predictive collection scheduling from 7-day fill velocity, dual-stakeholder platform (client dashboard + Transguard ops command center). AED 3.2M ARR from dashboard licensing. ClientPulse = six-dimension daily health scoring (SCDM uptime, collection SLA, ERP rate, tickets, QBR, AM contact frequency), RED/AMBER/GREEN banding, 5-step auto-playbook for RED accounts, 60-day renewal horizon pipeline, upsell trigger on health recovery. AED 4.2M single Emirates NBD renewal at risk. Key product insight: both products are zero new infrastructure — purely presentation + alerting layers on top of Transguard's existing operational data. Key proof mapping: PhonePe merchant dashboard (same: back-office data → client-facing intelligence layer); PhonePe milestone retention engine (same: data-triggered intervention playbooks). Resume title: "Manager — Products (Cash Services) & B2B Platform Growth". ATS keywords woven in: product lifecycle management, POC launch, client account management, new-to-company revenue, customer satisfaction, cross-functional delivery, vendor analysis, market research. Portfolio slug: rw-transguard. |
| 2026-05 | Nykaa | PM-2 Superstore (Bangalore, India) | QuickStore + AdAgent + SellerIQ | Portfolio shipped | India job — NO relocation note. Palette: pink #FC2779 (QuickStore) / purple #8B5CF6 (AdAgent) / gold #F59E0B (SellerIQ). All 3 user-selected from B2B-space options: QuickStore = 30-sec store-type wizard + peer-group smart shelf + pre-filled reorder cart → eliminates BDE calls; AdAgent = goal input → 5-step agent execution → 6-hr optimisation loop → plain-language weekly report → unlocks 60% never-advertised sellers; SellerIQ = daily AI brief (3 root causes) + NL Q&A + predictive alerts (stockout/return spike 48-72hr advance) + competitive benchmarking. Three-card portfolio block (rw-nykaa), 3 decks (8 slides each), strict one-page resume (8.3pt, 13px 25px padding). Merchant info correctly stated throughout. Resume title: "Product Manager — B2B Marketplace, Seller Growth & AI-Powered Commerce". |
| 2026-05 | Revolut | Junior PMM — Loyalty (UK, remote-eligible) | GrowthIQ + PlanPass | Portfolio shipped | Not UAE, not India — no relocation note. Palette: #0D0F14 dark / #0668E1 blue (GrowthIQ accent) / #7C3AED purple (PlanPass accent) / #F59E0B gold / #10B981 green. GrowthIQ = CDP propensity engine for plan upgrade: HIGH (>0.75) no offer, MED (0.55–0.74) minimum calibrated offer, LOW (<0.55) suppress; story variant selection (Aspiration / Fear / Social Proof) per band; Fear variant drove 31% CVR vs Aspiration 24% vs Social 19% in A/B; £1.02M saved vs blanket spend. PlanPass = subscription trust-transfer referral: Premium user gifts a named 30-day plan trial to a specific friend; named gift screen shows £ value breakdown (FX £18 + Cashback £8 + Lounge £28 + Insurance £12 = £66 total); Day-25 ValueStatement nudge shows personalised £ earned so far at peak trial conversion moment; referrer earns 500 RevPoints on conversion; 3.2× trial-to-paid vs generic app referral. Key design note: GrowthIQ = blue (#0668E1) accent; PlanPass = purple (#7C3AED) accent — differentiated palette within same dark base. Consumer screens = light/white (.light class on phone div); ops/dashboard screens = dark. Case study separation rule strictly applied: experience bullets describe PhonePe work only, no GrowthIQ/PlanPass mentions; professional summary ends with "Built two Revolut-specific case studies: GrowthIQ (propensity-scored plan upgrade engine) and PlanPass (subscription trust-transfer referral)" — name + 5-word description only. Merchant info correctly stated: Instant Discount & EMI subvention (not KYC), 30-min TAT = offer ops not merchant onboarding. Resume title: "Product Manager — Loyalty, Subscription Growth & Marketing Intelligence". |
| 2026-06 | HealthifyMe | Senior PM — B2C AI Products (Bangalore, India) | UpgradeIQ + RetentionOS + MealMatch | Portfolio shipped | India job — NO relocation note. Palette: dark #051A0A / green #00C853 (UpgradeIQ) / red #F44336 (RetentionOS) / orange #FF7043 (MealMatch) / gold #FFB300. UpgradeIQ = propensity scoring (goal proximity × streak × engagement velocity) → HIGH/MED/LOW tier → contextual upgrade card only at HIGH (34% CVR vs 16% generic, no discount). Milestone badge (5kg Challenge) as commitment device on checkout. Annual plan pre-selected. RetentionOS = daily churn scoring per subscriber → AT-RISK flagged at day 8 (22 days before renewal) → intervention type matched to churn reason: plateau → Milestone Challenge (34% recovery, ₹0 cost), streak break → streak recovery, feature drop → feature unlock. Browser-frame prototype with live intervention type selector (changes message preview + recovery rate), A/B toggle, and playbook. MealMatch = regional cuisine preference (inferred from log history) × macro budget × meal time → 3 South Indian dinner suggestions → "Why MealMatch chose this" transparency → 1-tap log → weekly adherence ring (82%). Orange (#FF7043) accent differentiates from UpgradeIQ green. 3-card portfolio block (rw-healthifyme). Resume title: "Senior Product Manager — AI-Powered Consumer Products · Subscription Growth · Personalisation & ML Product Management". Key proof parallels: UpgradeIQ = Propensity-to-Transact ML (−32% burn); RetentionOS = ops intelligence layer (TAT 2d→4h, −68% tickets); MealMatch = ML marketplace segmentation (11% YoY revenue, 500+ partners). |
| 2026-06 | Tesco | PM — Online Payments API (Bengaluru, India) | WalletBridge + AuthSense | Portfolio shipped | India job — NO relocation note. Palette: dark #0A1628 / blue #00539F (WalletBridge accent) / red #E2001A (AuthSense accent) / gold #F59E0B / muted #6B7A99. WalletBridge = self-serve payment method integration platform: 4-step wizard (API credentials → webhook config → mandatory 5-scenario sandbox gate → go-live review), method library showing active/pending/available methods, integration health dashboard with per-method alert thresholds. Maps directly to PhonePe self-serve PG integration platform (5K+ merchants, −23% CAC, zero eng dependency). AuthSense = real-time authorization intelligence console: 3σ anomaly detection per PSP × method × issuer, 24h auth rate trend chart, alert banner (clickable → alert detail), issuer heatmap (8 issuers × 5 methods, color-coded), remediation console with 5-step playbook. Maps directly to PhonePe unified ops intelligence layer (TAT 2d→4h, −68% escalation tickets). Both browser-frame prototypes. JD explicitly called out "optimising authorisation success rates and payment retry strategies" — AuthSense is that initiative as a ready-to-build product. Resume title: "Product Manager — Payments API · Payment Platform Design · Authorization Intelligence · Checkout Optimisation". Skills: Payments API Product Management, Authorization Success Rate Optimisation, Multi-Method Payment Integration, Platform & Self-Serve Design. |
| 2026-06 | Meesho | PM-2 Growth, Checkout & Returns Platform (Bangalore, India) | ReturnSense + Vaani-Cart + TrustBuy | Portfolio shipped | India job — NO relocation note. Palette: pink #F43397 (primary) / dark #353543 (secondary). ReturnSense = ML-driven returns mitigation at checkout (warnings, size matching, COD overlays with WhatsApp voice confirmation via Vaani), reducing returns by 40% and COD RTO by 25%. Vaani-Cart = voice-first checkout converting spoken address landmarks to coordinates, providing 14.2% lift. TrustBuy = supplier quality incentives linking SKU return rates to PRISM organic search visibility. All 3 prototypes built separately. Resume title: "Product Manager 2 — Growth, Checkout & Returns Platform". Skills: growth checkout, returns mitigation, RTO reduction, PRISM AI, Vaani voice assistant. |
| 2026-06 | Almosafer | Product Manager — Growth, Conversion & Gamification (Dubai, UAE) | SmartLock + RouteIQ + Milestone Explorer | Portfolio shipped | UAE job — relocation note included. Palette: dark #07132B / orange #FF5A00 (SmartLock) / green #10B981 (RouteIQ) / gold #F5A623 (Milestone Explorer). SmartLock = dynamic price-freeze engine: locks flight/hotel rates for 48 hours for a micro-fee, using dynamic volatility hedging API to protect margins (+22% CVR lift). RouteIQ = dynamic cart bundling query-routing taxi (Careem) and attraction APIs with real-time discounts, featuring 30-min self-serve partner campaign tools (+35% AOV, +25% attach). Milestone Explorer = gamified regional travel trail mapping destination-based steps (inspired by Toss Korea), utilizing user-uploaded photos of completed trips inside node circles to form visual journals (+17% CLTV, 2.4x shares). All 3 prototypes built separately. Resume title: "Product Manager — Growth, Conversion & Gamification". |
| 2026-06 | eyewa | PM, Engagement · Consumer Experience (Dubai, UAE) | ProLoop + RepeatIQ | Portfolio shipped | UAE job — relocation note included. Palette: ProLoop teal #00BFA5 / dark #0A1F1C; RepeatIQ indigo #6366F1 / dark #0F0E1A. Case studies pre-discussed with Ajay before build: 5 concepts brainstormed, user confirmed "Build A+B together (ProLoop) & C separately (RepeatIQ)". Both added to EXISTING rw-eyewa section (not a new company block — eyewa already had 3 ops case studies). eyewa app design research done before prototype build: light-mode app, white backgrounds, teal #00BFA5, minimal e-commerce aesthetic, bottom nav, category pills. ProLoop (6 screens): Home LensLoop alert (12-day countdown, red progress bar) → LensLoop supply detail (OD/OS wear grid, pre-filled address + card) → Confirm reorder → Order confirmed (confetti) → Pro dashboard (AED 847 hero, 8.5× ROI badge, savings breakdown BOGO/shipping/warranty) → Cancel-save flow (math argument: AED 847 / AED 99 = AED 748 net gain, Pause 30 days as primary CTA). RepeatIQ (5 screens): Before (generic home, no personalization, annotated with "what's missing") → With RepeatIQ (personalized "For You" tray: lens countdown + cross-category intro + Rx check reminder) → Complete Your Eyecare (68% Rx glasses users cohort, AED 29 trial, Rx on file, 78% convert within 14 days) → Win-Back (82-day lapse, saved cart, unused Pro BOGO, comeback offer) → Activated (confetti, Complete Eyecare milestone badge, next steps auto-queued, outcome stats). Key proof-point: same milestone-waiver psychology from PhonePe A/B (60% activation lift) applied to Pro cancel-save and renewal. Same ML propensity model from PhonePe (32% efficiency gain) applied to RepeatIQ cohort scoring. 8-digit hex alpha codes (e.g. #00BFA544) warn in pptxgenjs but decks still generate — border colors fall back to black, cosmetic only. Resume title: "Product Manager · Consumer Engagement & Behavioural Personalisation". Skills emphasise: repeat purchase engines, subscription retention, behavioural personalisation, cross-category LTV, win-back flows. |
| 2026-07 | Nomupay | Senior PM, Merchant Experience (Dubai, UAE) | MerchantIQ + SupportIQ + BatchPayIQ | Portfolio shipped | UAE job — relocation note: "Planning active relocation to Dubai, UAE in Q3 2026 · Requires visa sponsorship". Palette: #060C1A dark / #0F1F3D hero / #2563EB accent / #F59E0B gold / #059669 green. Two separate case studies: A+B combined as MerchantIQ (unified multi-acquirer analytics portal + SupportIQ embedded AI assistant, 74% ticket deflection, 8-sec resolution) + C standalone as BatchPayIQ (batch payout with pre-submission IBAN validation, real-time per-row status, 97.8% success rate, auto-reconciliation). Both prototypes 5 screens, single centered phone layout, sourcing sections with 4 industry citation cards each (Zendesk, Intercom, McKinsey, PYMNTS for MerchantIQ; Bottomline, Kyriba, Gartner, SWIFT for BatchPayIQ). GLOBAL RULE ADDED: every projected metric in prototype + deck must cite a named industry source in a sourcing section — rule 5 in Global Agent Rules. Resume title: "Senior Product Manager · Payments Infrastructure, Merchant Experience & B2B Platforms". Skills: Merchant Portal PM, Self-Service Platform Design, Multi-Acquirer Intelligence, Batch Payout Processing, Conversational AI Products. Case studies in resume: MerchantIQ, BatchPayIQ, J.P. Morgan PayIQ, Tesco WalletBridge. |

### Improvements To-Do
- [x] Build a reusable HTML resume base template — CSS baseline now locked in the "HTML Resume CSS baseline" section above
- [ ] Add a "Prototype preview image" system — auto-generate a screenshot thumbnail for each new prototype to use in the portfolio card (currently using text mockups in the deck)
- [ ] Create a standard prototype `base.html` starter file with all shared CSS (phone chrome, progress bars, confetti, fonts) to avoid copy-pasting boilerplate
- [ ] Add LinkedIn profile URL once confirmed

---

## QUICK COMMAND REFERENCE

When Ajay says `/apply [Company] — [JD text or link]`:

```
Phase 1 → Parse JD → Brainstorm 3–5 case studies → Get input → Confirm brief
Phase 2 → Build HTML prototype (4 screens) → QA checklist
Phase 3 → Build PptxGenJS deck (8 slides) → Extract text QA
Phase 4 → Add portfolio block → git push → Update case studies table in this file
Phase 5 → Write HTML resume → Run Chrome headless → PDF auto-generated → open PDF
Phase 6 → Write LinkedIn note (≤300 chars) + email → QA
```

**Estimated effort per run:** 45–90 minutes (prototype is the longest phase)
**Deliverables:** 1 prototype HTML, 1 PPTX deck, 1 HTML resume, **1 PDF resume** (auto-generated), 1 outreach HTML — all in one session. Phase 5 is not complete without the PDF.
