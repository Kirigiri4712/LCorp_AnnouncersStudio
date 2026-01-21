# Announcers Studio

**English** | [日本語](README_ja.md)

<img width="1196" height="863" alt="スクリーンショット 2026-01-21 225507" src="https://github.com/user-attachments/assets/04ef96e0-42c8-4ebf-97f4-69ce4f7476cf" />
VOICEVOX:Zundamon


This is not a Lobotomy Corporation MOD itself, but a tool to help create announcers for "L Corp's Announcers" MOD.
It can be used regardless of language.

Below is the Quick Start section copied from the readme.

--------------------------------------------------------------------------------
                            Quick Start
--------------------------------------------------------------------------------
1. Launch AnnouncerTool.exe
2. Click the "Add" button next to "Language:" to add a language code (e.g., "en", "jp")
3. Click each tag in the list to edit the dialogue
4. Assign Announcer.png (512x512) and UI.png (345x213)
5. Select a save folder and click "Save Mod"
6. Copy the generated folder to Lobotomy Corporation's MOD folder

For detailed usage and other features, please refer to the included readme.txt (English/Japanese).

Let's create your favorite announcers!

Thanks to LikeEatBanana for creating this mod!

## Download

**No build required for general users.** Download the pre-built executable from the [Releases](../../releases) page.

## Usage (GUI)

1. Run the downloaded `AnnouncerTool.exe`.

2. What you can do with the app:
   - Select an output folder with `Select Save Folder...` (folders like `AnnouncersImage_LEB` will be created here).
   - Add an announcer with `Add` and select it from the list.
   - Choose a language and enter dialogue for each tag (e.g., `AgentDie_TEXT`).
   - Use `Assign Image...` to assign a PNG to a tag (it will be copied when saving).
   - Export to the specified folder with `Save Mod`.
   - Load an existing `AnnouncersXML_LEB.xml` with `Load Existing` to reference existing text and images.

Notes:
- Audio files are not auto-generated (the game simply won't play audio if files are missing).
- When saving, placeholder images are automatically generated if `Announcer.png` or `UI.png` don't exist for each announcer.

For detailed instructions, refer to the included `HowToUse_en.txt`.

---

## Developer Information

### Branch Strategy

All development is done directly on the `main` branch (feature branches are not used).

### Development Tools

This project is developed using [Claude Code](https://claude.ai/code).

### Requirements

- .NET 8.0 SDK
- Windows (uses Windows Forms)

### Build

```bash
cd AnnouncerTool
dotnet build -c Release
```

### Run

```bash
dotnet run --project AnnouncerTool
```
