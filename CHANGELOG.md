# Changelog

All notable changes to PatentTools for Microsoft Word are documented in this file.

The project uses semantic-style versioning while it remains in pre-1.0 development.

## [0.2.0-beta]
### Added
- Populate list of reference signs based on description. Go to "Edit reference sign list", click on "Populate".
- Verification of API connection
- Auto-detection of available models
- Table of reference signs is now persistent on a document level.
- Call model in streaming mode for better progress feedback.
- System prompt for reference sign insertion is now mostly user-editable.

### Changed
- Separated editing the reference sign table and inserting reference into two separate ribbon icons.
- Reworked the settings dialog.
- Heavy editing on the README, removed fluff, adapted to latest version.
- Added a bogus description to the already present bogus claims for testing the population feature.

## Fixed

- Unicode processing should now work on Mac

### TBD

- Try to make it work with other models
- Updated Screenshots etc.
- Export and Import settings.

## [0.1.1] - 2026-08-17

### Added
- About button in the Ribbon Configuration group.
- Confirmation prompt before processing an entire document when no claim text is selected to avoid unnecessary token burning.
- Support for Ollama’s OpenAI-compatible local endpoint.
- Improved paragraph re-alignment for model output containing extra or omitted near-empty paragraphs.
- Improved Unicode-aware word-token handling for claim alignment.

### Changed

- Reorganized the Github repo, it now contains source code and a build script, whereas the ready .dotm file along with README and CHANGELOG are distributed in ZIP format as Github releases.

### Fixed
- Fixed an issue in which a reference sign could be inserted after the preceding word when source or model text contained unusual  Unicode characters.
- Fixed persistence of dot-decimal temperature values on locales using a comma decimal separator.

## [0.1.0] - 2026-06-30

### Added
- Initial public release of Patent Tools for Microsoft Word.
- Reference-sign insertion using an OpenAI-compatible chat-completions endpoint.
- Track Changes markup for inserted reference signs.
- Custom Ribbon tab and Settings dialog.
- Persistent per-user settings.