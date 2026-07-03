"""SMRC Key Personnel (KP) screening criteria.

Transcribed from the eight Commence KP Screening & Competitive Comparison
checklists (v2.0, Source: SMRC SOW, June 2026) in
"06 Staffing and KPs/Claude Internal Resume Review/KP Screening Checklists".

Each role has two parts, matching the checklists:
  - knockouts: minimum requirements. A candidate must clearly meet ALL of
    these to advance for that role. Missing/ambiguous evidence = not met.
  - differentiators: strategic, above-minimum qualities used to rank
    qualified candidates against each other and the incumbent profile.
    These do NOT gate a candidate out — they inform fit_strength.

If the SOW or a checklist changes, edit the relevant block below and
re-run smrc_screen.py. Nothing else needs to change.
"""

KP_CRITERIA = {
    "PM": {
        "title": "Project Manager",
        "knockouts": [
            "Bachelor's (healthcare-relevant) + Master's in Accounting, Business, Public Admin, "
            "or healthcare-relevant field (or Bachelor's + 4 additional years of related experience)",
            "Current PMP certification, held and maintained (required per June 2026 SOW)",
            "Available as a fully dedicated 1.0 FTE to the contract",
            "5+ years professional experience (healthcare, federal, or commercial) as PM, Project "
            "Leader, Operations Manager, Technical PM, or Product Manager",
            "3+ years managing oversight of complex medical review systems & workflow (beyond "
            "personnel management)",
            "3+ years Medicare Fee-for-Service (FFS) experience — coverage & payment rules",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Direct CMS program experience (agency-side or as a CMS contractor)",
            "CMS Center for Program Integrity (CPI) experience — SMRC, UPIC, Recovery Audit, or MAC "
            "program-integrity work",
            "Prior Supplemental Medical Review Contractor (SMRC) experience (incumbent-equivalent)",
            "Quality assurance program ownership (95% accuracy standard, inter-rater reliability)",
            "Managed high-volume national medical review operations (Task 1 PCG + Task 2 PI scale)",
            "Deliverable & reporting cadence management (MSR, monthly cost reports, dashboard/metrics)",
            "Contract ramp-up/transition/close-out leadership (90-day ramp, Task 13)",
            "CMS COR & stakeholder engagement; weekly status meeting facilitation",
        ],
    },
    "APM": {
        "title": "Assistant Project Manager",
        "knockouts": [
            "Bachelor's (healthcare-relevant) + Master's in Accounting, Business, Public Admin, "
            "or healthcare-relevant field (or Bachelor's + 4 additional years of related experience)",
            "PMP certification held, or actively working toward PMP (required per June 2026 SOW)",
            "Fully dedicated 1.0 FTE; able to function as PM in the PM's absence",
            "3+ years professional experience (healthcare, federal, or commercial) as PM, Project "
            "Leader, Operations Manager, Technical PM, or Product Manager",
            "2+ years managing oversight of complex medical review systems & workflow (beyond "
            "personnel management)",
            "2+ years Medicare Fee-for-Service (FFS) experience — coverage & payment rules",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Direct CMS program experience (agency-side or as a CMS contractor)",
            "CMS CPI experience — SMRC, UPIC, Recovery Audit, or MAC program-integrity work",
            "Prior Supplemental Medical Review Contractor (SMRC) experience (incumbent-equivalent)",
            "Readiness to step up to acting PM (continuity/succession strength)",
            "Supported high-volume national medical review operations (Task 1 PCG + Task 2 PI)",
            "Project planning, scheduling, & deliverable tracking against CMS timelines",
            "Reporting support (MSR, cost reports, dashboard/metrics) & issue/risk logs",
            "CMS COR/status-meeting support & minutes ownership",
        ],
    },
    "CMD": {
        "title": "Contractor Medical Director",
        "knockouts": [
            "Active MD or DO license, currently licensed to practice in the U.S.",
            "Board-certified",
            "CMS Administrative Law Judge (ALJ)-eligible",
            "3+ years practicing as an MD/DO (currently licensed)",
            "2+ years in health insurance/utilization review/claims processing (coverage/"
            "medical-necessity policy role)",
            "Extensive Medicare Fee-for-Service (FFS) knowledge",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Direct CMS program experience (agency-side or as a CMS contractor)",
            "CMS Center for Program Integrity (CPI) experience — SMRC, UPIC, Recovery Audit, or MAC "
            "program-integrity work",
            "Prior Supplemental Medical Review Contractor (SMRC) experience (incumbent-equivalent)",
            "Currently or previously CMS-approved as a Contractor Medical Director (de-risks the "
            "CMS approval gate)",
            "LCD/national coverage & medical-necessity policy development experience",
            "Statistical extrapolation/probability-sampling exposure (SMRC Optional Tasks 2 & 3)",
            "Fraud/program-integrity-focused clinical review (Task 2 PI/FIG alignment)",
            "High-volume, multi-specialty national medical review breadth (Task 1 PCG alignment)",
        ],
    },
    "MRM_PCG": {
        "title": "Medical Review Manager — Task 1 (PCG Specialty Reviews)",
        "knockouts": [
            "Active RN license (State, DC, Puerto Rico, or U.S. territory)",
            "Master's in nursing or related field (or BSN + 3 additional years of relevant experience)",
            "No prior Medicare sanction or program exclusion (verified)",
            "5+ years clinical experience (acute care, SNF, and/or office/clinic practice)",
            "5+ years medical review experience, including 3+ years management",
            "Extensive Medicare Fee-for-Service (FFS) knowledge — coverage & payment rules",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Direct CMS program experience (agency-side or as a CMS contractor)",
            "CMS Center for Program Integrity (CPI) experience — SMRC, UPIC, Recovery Audit, or MAC "
            "program-integrity work",
            "Prior Supplemental Medical Review Contractor (SMRC) experience (incumbent-equivalent)",
            "High-volume, multi-specialty national post-pay medical review (Task 1 PCG alignment)",
            "Certified professional coder (CPC) or coding-determination oversight experience",
            "LCD/NCD coverage & medical-necessity policy application across MAC jurisdictions",
            "Inter-rater reliability/QA program design for clinical reviewers (95% accuracy standard)",
            "Discussion & Education (D&E) and provider/supplier engagement experience (Task 4)",
        ],
    },
    "MRM_PI": {
        "title": "Medical Review Manager — Task 2 (Program Integrity)",
        "knockouts": [
            "Active RN license (State, DC, Puerto Rico, or U.S. territory)",
            "Master's in nursing or related field (or BSN + 3 additional years of relevant experience)",
            "No prior Medicare sanction or program exclusion (verified)",
            "5+ years clinical experience + 5+ years medical review (including 3+ years management)",
            "2+ years fraud-focused medical review experience (REQUIRED — Task 2 hard gate)",
            "Extensive Medicare Fee-for-Service (FFS) knowledge — coverage & payment rules",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Fraud Investigations Group (FIG)/FWA-focused review experience",
            "CMS Center for Program Integrity (CPI) experience — SMRC, UPIC, Recovery Audit, or MAC "
            "program-integrity work",
            "Prior Supplemental Medical Review Contractor (SMRC) experience (incumbent-equivalent)",
            "UPIC/NBI MEDIC/ZPIC or law-enforcement coordination experience",
            "Detecting record falsification/alteration & aberrant billing pattern analysis",
            "Statistical extrapolation/probability-sampling exposure (Optional Tasks 2 & 3)",
            "Vulnerability identification & program-integrity lead development",
            "Provider/supplier-specific & data-pattern-driven medical review (Task 2 scope)",
        ],
    },
    "CIO": {
        "title": "Chief Information Officer",
        "knockouts": [
            "Bachelor's in info systems, computer science, or related tech field (relevant "
            "experience may substitute)",
            "Available as a fully dedicated 1.0 FTE (cannot be combined with the SSO role)",
            "Medicare Fee-for-Service (FFS) program knowledge",
            "5+ years combined IT work experience",
            "3+ years healthcare IT (federal or commercial) as CIO, IT manager, CTO, or network "
            "administrator",
            "Experience documenting & implementing IT policy aligned to enterprise security "
            "requirements",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Direct CMS program experience (agency-side or as a CMS contractor)",
            "CMS CPI/SMRC, UPIC, Recovery Audit, or MAC IT-compliance experience",
            "Led a FedRAMP-compliant ATO/agency Authority to Operate (ATO) effort",
            "FISMA, NIST SP 800-53, & CMS BPSSM/ARS/IS2P2 compliance leadership",
            "CMS Security Assessment & Authorization (SA&A)/CFACTS package ownership",
            "Cloud security & CMS Cloud Services/FedRAMP 3PAO assessment experience",
            "IT governance: investment management, SDLC/TLC, configuration & asset management",
            "Stood up secure IT environment under a 90-day ramp-up/contract transition",
        ],
    },
    "SSO": {
        "title": "Systems Security Officer",
        "knockouts": [
            "Bachelor's in info systems, computer science, or related tech field (relevant "
            "experience may substitute)",
            "Independent of IT operations (no IT ops, maintenance, or development responsibility)",
            "Available as a fully dedicated 1.0 FTE; Medicare program knowledge",
            "5+ years combined work experience",
            "3+ years healthcare IT security as SSO, IT specialist, security engineer, info "
            "security analyst, or info systems technician",
            "40+ hrs/yr continuing professional education (CPE) from a recognized IS security org "
            "(CSCOUT counts)",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Direct CMS program experience (agency-side or as a CMS contractor)",
            "CMS CPI/SMRC, UPIC, Recovery Audit, or MAC information-security experience",
            "FedRAMP/FISMA continuous monitoring & POA&M management",
            "Incident & breach response (1-hr reporting, CMS IRT/US-CERT coordination)",
            "BPSSM (IOM 100-17)/ARS/IS2P2 control implementation & oversight",
            "Vulnerability management & SCAP-compliant scanning/patch remediation "
            "(15/30/90/360-day)",
            "Security clearance designation, access provisioning, & user-account auditing",
            "Security certification (CISSP, CISM, CAP, or equivalent)",
        ],
    },
    "CSM": {
        "title": "Customer Service Manager",
        "knockouts": [
            "Bachelor's degree (accredited) + customer-service-related training (or 4+ years "
            "related work experience in lieu of degree)",
            "Knowledgeable about the Medicare program; able to resolve escalated concerns",
            "Available as a fully dedicated resource to the contract",
            "2+ years as manager of a customer service unit",
            "Track record handling escalations — issues unresolved by CS staff routed to this role",
            "Able to staff/operate toll-free support, 8:30 a.m.-6:00 p.m. ET (Task 10)",
            "No prior sanctioning by any licensing/regulatory/professional body",
            "Not excluded or debarred from Medicare/federal programs (OIG LEIE + SAM.gov)",
        ],
        "differentiators": [
            "Medicare provider/supplier inquiry handling & provider-education experience (the key "
            "credential)",
            "CMS CPI/SMRC, UPIC, Recovery Audit, or MAC customer-service experience",
            "Prior Supplemental Medical Review Contractor (SMRC) experience (incumbent-equivalent)",
            "Internal Commence candidate (BFCC-QIO call-center leadership) — reinforces "
            "systems/process story, avoids external sourcing lead time",
            "Built/managed provider FAQ, scripts, & secure mailbox/correspondence triage",
            "Service metrics & QA: 2-day acknowledgment, 14-day update cadence, call "
            "monitoring/recording",
            "ADR provider-outreach & reminder-call operations to drive documentation response",
            "Ad-hoc reporting & report-of-contact logs delivered to CMS on request (48-hr)",
        ],
    },
}

# Resumes showing current employment at any of these get an OCI flag for
# Bryan/Albie to review. Not exhaustive — the model is also instructed to
# flag any other apparent CMS medical-review competitor (UPIC, Recovery
# Audit Contractor, MAC, QIO) it recognizes from resume text.
OCI_WATCHLIST = [
    "Noridian Healthcare Solutions (current SMRC incumbent)",
]
