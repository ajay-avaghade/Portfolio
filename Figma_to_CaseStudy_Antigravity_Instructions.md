# Antigravity Workflow: Automated Figma Screen Extraction & Semantic Integration

## 🎯 Primary Objective
Automate the extraction of high-resolution mobile and web screens from a provided Figma file, analyze their visual and textual content, and seamlessly integrate them into the correct `[Insert Design: ...]` placeholders within a given Markdown case study.

## 🛠️ Prerequisites & Configuration
To execute this workflow, the following environment variables and inputs must be provided to the Antigravity agent:
* `FIGMA_PERSONAL_ACCESS_TOKEN`: Valid Figma API token.
* `FIGMA_FILE_KEY`: The 22-character alphanumeric key from the Figma file URL.
* `CASE_STUDY_MD_PATH`: File path to the target markdown case study.
* `EXPORT_DIRECTORY`: Destination folder for high-resolution images.

---

## 🤖 Agent Execution Steps

### Step 1: Parse the Target Case Study
1. Read the markdown file located at `CASE_STUDY_MD_PATH`.
2. Extract all image placeholders. Look for the specific pattern: `*[Insert Design: <Description>]*` or similar descriptive tags.
3. Create a **Placeholder Mapping Table** storing:
    * Placeholder ID
    * Contextual Description (e.g., "Homepage category tiles with dynamic offer badges")
    * Surrounding Text (to provide semantic context for matching)

### Step 2: Fetch & Analyze Figma Document Tree
1. Call the Figma API `GET /v1/files/:key` to retrieve the document structure.
2. Traverse the node tree to identify all `FRAME`, `COMPONENT`, or `INSTANCE` nodes that represent complete screens. 
    * *Heuristic:* Look for top-level frames on the canvas or nodes with names like "Screen", "Mockup", "Homepage", "Checkout", etc. Ignore small nested elements like icons or buttons unless explicitly requested.
3. Extract the metadata for each identified screen:
    * `node_id`
    * `name`
    * `text_content` (Concatenate all child `TEXT` nodes to understand what is on the screen).

### Step 3: Semantic Matching Engine
1. Use your LLM capabilities to compare the **Placeholder Mapping Table** (from Step 1) against the **Figma Screen Metadata** (from Step 2).
2. For each placeholder description, calculate the best match among the available Figma screens based on:
    * Figma frame names.
    * Text content inside the Figma frame (e.g., if the placeholder asks for "Checkout nudges", look for frames containing text like "Change payment method", "Offers", "Checkout").
3. Generate a confidence score for each match. If the confidence score is below 70%, flag it for manual human review.

### Step 4: High-Resolution Extraction
1. For all successfully matched `node_id`s, construct a comma-separated string.
2. Call the Figma API `GET /v1/images/:key?ids=<node_ids>&scale=4&format=png` (Use `scale=4` to ensure the absolute highest resolution for Retina displays and infographics).
3. Download the images from the returned AWS S3 URLs.
4. Save the downloaded images to the `EXPORT_DIRECTORY` using semantic filenames (e.g., `01_homepage_discovery.png`, `02_checkout_nudges.png`).

### Step 5: Markdown Synthesis & Infographic Formatting
1. Re-read the original case study markdown.
2. Replace the identified placeholders with standard markdown image syntax pointing to the local `EXPORT_DIRECTORY` paths.
    * *Format:* `![<Figma Frame Name>](./<EXPORT_DIRECTORY>/<filename>.png)`
3. **Infographic Augmentation:** If the case study text implies a flow between two screens, wrap the images in an HTML table or a flex-box div (if supported by the target renderer) to display them side-by-side for comparison.
4. Output the final, modified markdown as a new file: `<original-filename>_visual_final.md`.

---

## 🚨 Error Handling & Guardrails
* **Rate Limits:** Implement exponential backoff for Figma API calls to avoid 429 Too Many Requests errors.
* **Missing Screens:** If a placeholder describes a design that does not exist in the Figma file, leave the placeholder intact and append an inline warning: `<!-- WARNING: No matching Figma screen found for this description -->`.
* **Large Files:** If the Figma file contains hundreds of screens, restrict the initial `GET /v1/files/` search to specific pages (e.g., "Final Designs", "Handoff") if defined in the prompt.
