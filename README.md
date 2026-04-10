# Handl Offer Letter Generator - Office Add-in

This is a Word Online taskpane add-in that enables users to generate customized offer letters directly within Microsoft Word. The add-in provides a streamlined interface for creating professional employment offer documents with pre-populated fields and templates.

## Setup

1. **Clone the repository** and navigate to the project directory:
   ```bash
   git clone <repository-url>
   cd offer-letter-addin
   ```

2. **Update manifest.xml URLs** to point to your GitHub Pages deployment URL. Replace the placeholder URLs in `src/manifest.xml` with your GitHub Pages URL (e.g., `https://yourusername.github.io/offer-letter-addin/`).

3. **Sideload the manifest in Word Online**:
   - Open Word Online (office.com)
   - Go to Insert > Get Add-ins > My Add-ins > Upload My Add-in
   - Upload the `src/manifest.xml` file
   - The add-in will appear in your Word taskpane

## Deployment

The `deploy.yml` GitHub Actions workflow automatically deploys the contents of the `src/` folder to GitHub Pages on every push to the `main` branch.

## Organization-Wide Deployment

For org-wide deployment across your organization, see the System Specification document.
