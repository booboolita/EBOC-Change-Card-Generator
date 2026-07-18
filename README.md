# EBOC Change Card Generator

## Project Overview

The EBOC Change Card Generator is a prototype web application designed to support evidence-based guideline revision work for the Evidence-Based Outcomes Center (EBOC).

The tool helps convert a focused clinical evidence question, current guideline materials, workgroup notes, and literature evidence into a physician-facing Change Card for guideline review.

The purpose of the tool is to make guideline revision more efficient, structured, and evidence-based while keeping final clinical approval human-led.

---

## Current Prototype Status

This project has completed an initial proof of concept using Google Colab.

The proof of concept successfully generated physician-facing Change Cards for three Bronchiolitis guideline update questions:

1. High-flow nasal cannula initiation in bronchiolitis
2. Albuterol use in bronchiolitis
3. Feeding, NG feeds, and IV fluids in bronchiolitis

The next phase is to convert the working Colab process into a simple browser-based web app so users do not need to interact with Colab code.

---

## Intended Users

This tool is intended for EBOC guideline revision workflows, including:

- EBOC project managers
- Guideline revision workgroups
- Physician reviewers
- EBOC medical leadership
- Clinicians participating in evidence-based pathway and guideline updates

---

## What the Tool Does

The tool accepts guideline revision inputs and generates a physician-facing Change Card.

### Inputs

The user provides:

- Focused topic
- Evidence question
- Guideline last updated date
- Current guideline draft
- Current clinical algorithm or pathway
- Workgroup notes
- Literature evidence
- Optional local pathway context

### Processing

The tool:

1. Checks that required fields and documents are present.
2. Reads uploaded Word and PDF documents.
3. Extracts relevant text from the documents.
4. Combines the current guideline context, workgroup concern, and literature evidence.
5. Applies the physician-facing Change Card prompt.
6. Generates a structured Change Card.
7. Creates a downloadable Word document.

### Outputs

The tool produces:

- Physician-facing Change Card preview
- Downloadable Word document
- Evidence-based proposed guideline edit or new section
- GRADE-informed confidence rating
- Evidence traceback notes with citations when available

---

## Physician-Facing Change Card Sections

The generated card uses this structure:

1. Evidence Question
2. Current Guidance
3. Workgroup Concern / Reason for Review
4. Evidence Finding
5. Proposed Guideline Edit or New Section
6. GRADE-informed Confidence Rating
7. Decision Required Before Approval
8. Evidence Traceback Notes

---

## Section 5 Placement Guidance Requirement

Section 5 should not only provide proposed wording. It should also help the EBOC team understand where the proposed text belongs in the guideline.

Section 5 should include:

- Recommended update type
- Suggested placement in the current guideline
- Rationale for placement
- Proposed guideline text
- Addendum recommendation, if applicable
- Related alignment notes

The tool should not force an addendum for every update.

An addendum should only be recommended when the proposed content is:

- Too long for the main guideline body
- Criteria-based
- Operational
- Stepwise
- Likely to clutter the main guideline
- Better suited as a reference section

When an addendum is not needed, the tool should recommend the simplest appropriate placement, such as replacing existing text, adding a clarifying sentence, or adding a short subsection.

---

## Word Document Formatting Goals

The downloadable Word document should require minimal cleanup before physician review.

The document should include:

- Clear date and topic header
- Clean title
- Numbered section headings
- Bolded key conclusions and labels
- Readable paragraph spacing
- Proper bulleted lists
- Proper numbered lists
- Complete Evidence Traceback citations with links when available
- Minimal Markdown leftovers

---

## Human-in-the-Loop Expectations

This tool is intended to support, not replace, clinical judgment.

Human involvement is required at several points:

### 1. Evidence Question Selection

The workgroup or EBOC project manager identifies the focused clinical question or update topic.

### 2. Literature Evidence Collection

The user gathers literature evidence from OpenEvidence or another approved evidence source.

The current version does not automatically search OpenEvidence.

### 3. Document Upload

The user uploads the current guideline, current pathway or algorithm, workgroup notes, and literature evidence.

### 4. Change Card Review

The generated card is reviewed by the EBOC project manager and clinical workgroup.

### 5. Physician Approval

Physicians review the proposed guideline edit or new section and choose to approve, modify, or reject it.

### 6. Final Guideline Update

Final clinical approval and publication remain human-led.

The tool does not independently publish guideline changes.

---

## Out of Scope for Version 1

The first web app version will not include:

- Automatic OpenEvidence searching
- Automatic guideline editing
- Automatic publication
- Algorithm or pathway editing
- Login or user management
- Database storage
- Approval workflow dashboard
- PHI processing

The first version is focused only on generating physician-facing Change Cards from user-provided materials.

---

## Security and Data Notes

The prototype is intended for non-PHI use only.

The initial Bronchiolitis test materials are public-facing guideline-related documents.

No patient-specific data, MRNs, charts, or PHI should be uploaded.

API keys should not be stored in the code or committed to GitHub. API keys should be managed through the hosting platform's secrets or environment variable system.

---

## Planned Version 1 App Features

The first web app should include:

- Browser-based interface
- File upload boxes
- Focused topic field
- Evidence question field
- Last updated date field
- Status checks for missing files
- API key connection check
- Generate Change Card button
- On-screen card preview
- Download Word document button
- Styled Word output

---

## Project Roadmap

### Phase 0: Proof of Concept

Completed in Google Colab.

Three Bronchiolitis physician-facing cards were generated and downloaded successfully.

### Phase 1: Simple Web App

Convert the Colab workflow into a browser-based Streamlit app.

### Phase 2: Improved Formatting and Usability

Improve Word formatting, status messages, and user experience.

### Phase 3: Saved Project Context

Allow repeated use of the same guideline/pathway/workgroup materials without re-uploading every time.

### Phase 4: Approval and Revision Tracking

Track generated cards, physician decisions, modifications, and final approved changes.

### Phase 5: Future Integrations

Potential future integrations may include evidence source integration, guideline document editing support, and approved-change tracking.

---

## Current Development Notes

This repository is currently being built as a prototype.

The immediate next development goal is to create a simple Streamlit app that reproduces the successful Colab workflow in a user-friendly web interface.
