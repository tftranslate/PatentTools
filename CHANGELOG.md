# Changelog

All notable changes to PatentTools for Microsoft Word are documented in this file.

The project uses semantic-style versioning while it remains in pre-1.0 development.

## [0.2.0-beta - work in progress - not yet released]
### Added
- Verification of API connection
- Auto-detection of available models
- Table of reference signs is now persistent on a document level.
- Call model in streaming mode for better progress feedback.
- System prompt for reference sign insertion is now mostly user-editable.

### Changed
- Separated editing the reference sign table and inserting reference into two separate ribbon icons.
- Reworked the settings dialog in preparation for further functionality.

## Fixed

- Unicode processing should now work on Mac

### To be done
- Implement auto-population of list of reference signs from description

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