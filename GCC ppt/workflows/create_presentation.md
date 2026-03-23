# Workflow: Create India GCC Landscape Presentation

## Objective
Build a 17-slide Canva presentation on the India GCC landscape for financial services, using research data from `research/` files.

## Prerequisites
- Research files completed: `research/01_location_comparison.md`, `research/02_gcc_vs_outsourcing.md`, `research/03_branding_evp.md`
- Sources consolidated: `research/04_sources.md`
- Canva MCP tools available

## Step 1: Review Research
Read all three research files to internalize key data points. Focus on:
- Specific numbers (not ranges) for each slide
- Source attribution for every claim
- Strategic recommendations that tie the narrative together

## Step 2: Submit Outline for Review
Use `request-outline-review` with:
- **topic**: "India GCC Landscape: Strategic Analysis for Financial Services"
- **audience**: "professional"
- **style**: "minimalist"
- **length**: "balanced"
- **pages**: 17 slides as defined in CLAUDE.md

### Slide Content Specifications

**Slide 1 — Title**
- Title: "India GCC Landscape: Strategic Analysis for Financial Services"
- Description: Location comparison, value proposition analysis, and branding strategy for financial services GCCs

**Slide 2 — India GCC Market Overview**
- 1,800+ GCCs employing 1.9M professionals generating $65B in revenue
- BFSI sector represents 35% of the GCC market with 230+ centers and 450K+ professionals
- Projected to reach 2,400+ GCCs and $100B revenue by 2030

**Slide 3 — Location Comparison: Four Hubs at a Glance**
- Side-by-side matrix covering GCC share, banking GCC share, cost index, office rent, attrition, talent pool, and weighted scorecard
- Bangalore 67/90, Hyderabad 66/90, Delhi NCR 64/90, Pune 58/90

**Slide 4 — Bangalore: The Established Leader**
- 870+ GCCs with 42% market share and 32% of banking GCCs
- Home to Goldman Sachs, JP Morgan, Wells Fargo, Citibank, Fidelity
- 1M+ tech professionals with India's first dedicated GCC Policy 2024-2029
- Challenge: Highest attrition at 16-20% and highest operating costs

**Slide 5 — Hyderabad: The Fastest Growing Challenger**
- 355+ GCCs capturing 20% of banking GCCs with fastest new GCC additions
- Office rents 20-30% lower than Bangalore with world-class HITEC City infrastructure
- Vanguard scaling from 300 to 2,300 employees by 2029
- TS-iPASS single-window clearance in 15 days

**Slide 6 — Pune: The High-Retention Engineering Hub**
- 250+ GCCs with best talent retention rates, 5-8 points better than Bangalore
- 15-20% lower total cost than Bangalore with 84% graduate employability
- Barclays operates its largest facility outside London here with 9,000 employees
- Challenge: No international flights and shallower BFSI specialist pool

**Slide 7 — Delhi NCR: Financial Services and Connectivity Powerhouse**
- 300+ GCCs across Gurugram Noida and Delhi
- Best international airport with 79.2M passengers and direct US UK Europe flights
- Best metro network at 394km with 289 stations
- Home to American Express, Deutsche Bank, Barclays, HSBC

**Slide 8 — Strategic Location Recommendation**
- Hub-and-spoke model with primary hub in Bangalore or Hyderabad
- 70%+ of banking GCCs operate multiple delivery centers for resilience
- Tier-2 satellite cities offer 30-40% cost savings and 10-12% lower attrition

**Slide 9 — GCC Model vs Third-Party Service Provider**
- Comparison across eight dimensions: upfront cost, long-term cost, control and IP, setup time, innovation, talent retention, regulatory compliance, and risk profile

**Slide 10 — GCC Model: Long-Term Strategic Value**
- Breakeven at Year 2-3, then 20% annual cost decline versus outsourcing
- 3.2x more digital patents per $1M invested
- 55% of enterprise tech products now originate in GCCs
- 63% of global CXOs say GCCs are central to innovation

**Slide 11 — Third-Party Provider: Speed vs Strategic Risk**
- 30% of all data breaches came from third parties in 2024
- Outsourcing failure costs routinely reach $0.5-1B per incident
- Average financial services data breach costs $6.08M
- EU DORA regulation targets outsourcing concentration risk

**Slide 12 — Recommended Operating Model**
- Hybrid approach with GCC for core strategic functions and selective outsourcing
- BOT model adoption surged from under 10% to 40% of new GCCs
- Companies with Mega GCCs outsource 2X more, showing hybrid is the future

**Slide 13 — GCC Branding and EVP: The Talent Challenge**
- 60% of GCC hiring comes from other GCCs creating intense circular competition
- Only 25-30% of Indian GCCs invest in employer branding
- AI/ML talent supply meets only 49% of demand with 1M+ shortage by 2027
- 75% of Gen Z employees intend to leave within 2 years

**Slide 14 — Building a Compelling EVP for Financial Services**
- Five pillars: compensation and long-term incentives, career and global mobility, innovation culture, work-life and hybrid flexibility, purpose and impact
- Gartner Human Deal framework increases EVP satisfaction by 15%
- 81% of GCCs deploy upskilling as retention and 95% adopted hybrid work

**Slide 15 — Strengthening Internal and External Branding**
- Internal: glocal identity, employee advocacy with 8X engagement, hackathons and innovation labs, stay interviews, Chief Storyteller role
- External: LinkedIn thought leadership, campus partnerships with IITs and NITs, industry events, Glassdoor and AmbitionBox management, Great Place to Work certification
- ROI: 43% lower cost-per-hire and 28% lower turnover and 50% more qualified applicants

**Slide 16 — Case Studies: Leading Financial Services GCCs**
- Goldman Sachs: 9,000+ in India, Most Admired GCC 2025, AI innovation lab
- JP Morgan: 50,000+ in India, $17B annual tech investment, Code for Good hackathon
- Barclays: 14,000+ in India, Pune largest outside London, DEI awards leader

**Slide 17 — Key Recommendations and Sources**
- Three strategic priorities: optimize multi-city location strategy, establish GCC with hybrid operating model, invest in differentiated EVP and employer brand
- Source references from Zinnov NASSCOM ANSR McKinsey EY JLL Gartner and others

## Step 3: Wait for User Approval
User reviews the outline in the Canva widget and either:
- Approves → proceed to Step 4
- Requests changes → update pages and call `request-outline-review` again

## Step 4: Generate Design
After approval, call `generate-design-structured` with:
- The approved outline parameters
- design_type: "presentation"
- All titles and descriptions with punctuation removed (Claude constraint)

## Step 5: Create Design from Candidate
- Review generated candidates
- Ask user to select preferred design
- Call `create-design-from-candidate` with selected job_id and candidate_id

## Step 6: Review and Export
- Use `start-editing-transaction` to review content accuracy
- Make any text corrections via `perform-editing-operations`
- Export as PDF/PPTX via `export-design`
- Share design link with user

## Quality Checklist
- [ ] All 17 slides present with correct content
- [ ] Every statistic has a source citation
- [ ] Data points match research files
- [ ] Consistent formatting and visual style
- [ ] Strategic narrative flows logically
- [ ] Recommendations are actionable and specific
- [ ] Source URLs are included on final slide

## Edge Cases & Notes
- Canva has 90-character limit per slide description (Claude constraint)
- Use "short" or "balanced" length only (not "comprehensive")
- Remove all punctuation from titles/descriptions before passing to generate-design-structured
- If design generation fails, retry with simplified descriptions
- Max 150 characters for topic field
