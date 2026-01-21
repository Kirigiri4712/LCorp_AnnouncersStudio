using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.VisualBasic;
using NAudio.Wave;

public class AppSettings
{
    public string Language { get; set; } = "en";
    public bool TranslatePlaceholders { get; set; } = true;
    public bool EnableUndoRedo { get; set; } = false;
    public bool NormalizeAudio { get; set; } = true;
    public int CompressionLevel { get; set; } = 50;
}

public static class Localization
{
    public static string CurrentLanguage { get; set; } = "en";

    private static readonly Dictionary<string, Dictionary<string, string>> Strings = new();
    private static readonly List<string> AvailableLanguages = new();

    static readonly string LanguagesFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Languages");

    public static void LoadLanguages()
    {
        Strings.Clear();
        AvailableLanguages.Clear();

        if (!Directory.Exists(LanguagesFolder))
        {
            Directory.CreateDirectory(LanguagesFolder);
        }

        var jsonFiles = Directory.GetFiles(LanguagesFolder, "*.json");
        foreach (var file in jsonFiles)
        {
            try
            {
                var langCode = Path.GetFileNameWithoutExtension(file);
                var json = File.ReadAllText(file);
                var strings = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                if (strings != null)
                {
                    Strings[langCode] = strings;
                    AvailableLanguages.Add(langCode);
                }
            }
            catch
            {
                // Skip invalid files
            }
        }

        // Ensure at least English exists
        if (!Strings.ContainsKey("en"))
        {
            Strings["en"] = new Dictionary<string, string>
            {
                ["LanguageName"] = "English",
                ["WindowTitle"] = "Announcer Creator"
            };
            if (!AvailableLanguages.Contains("en"))
                AvailableLanguages.Add("en");
        }
    }

    public static List<string> GetAvailableLanguages() => AvailableLanguages;

    public static string GetLanguageName(string langCode)
    {
        if (Strings.TryGetValue(langCode, out var langStrings) && langStrings.TryGetValue("LanguageName", out var name))
            return name;
        return langCode;
    }

    public static string Get(string key)
    {
        if (Strings.TryGetValue(CurrentLanguage, out var langStrings) && langStrings.TryGetValue(key, out var value))
            return value;
        if (Strings.TryGetValue("en", out var enStrings) && enStrings.TryGetValue(key, out var fallback))
            return fallback;
        return key;
    }

    public static string GetTagDisplayName(string internalTag)
    {
        // Extract base tag name (remove _1, _2 etc. suffix if present)
        var baseTag = internalTag;
        var suffix = "";
        var lastUnderscore = internalTag.LastIndexOf('_');
        if (lastUnderscore > 0 && int.TryParse(internalTag.Substring(lastUnderscore + 1), out _))
        {
            baseTag = internalTag.Substring(0, lastUnderscore);
            suffix = internalTag.Substring(lastUnderscore); // includes the underscore
        }

        var key = "Tag_" + baseTag;
        var displayName = Get(key);

        // If no translation found, return original
        if (displayName == key)
            return internalTag;

        return displayName + suffix;
    }

    public static string GetInternalTagName(string displayTag)
    {
        // Reverse lookup: find internal tag name from display name
        // Extract suffix if present
        var suffix = "";
        var lastUnderscore = displayTag.LastIndexOf('_');
        if (lastUnderscore > 0 && int.TryParse(displayTag.Substring(lastUnderscore + 1), out _))
        {
            suffix = displayTag.Substring(lastUnderscore);
            displayTag = displayTag.Substring(0, lastUnderscore);
        }

        // Search for the internal tag name
        if (Strings.TryGetValue(CurrentLanguage, out var langStrings))
        {
            foreach (var kvp in langStrings)
            {
                if (kvp.Key.StartsWith("Tag_") && kvp.Value == displayTag)
                {
                    return kvp.Key.Substring(4) + suffix; // Remove "Tag_" prefix
                }
            }
        }

        // Fallback to English
        if (Strings.TryGetValue("en", out var enStrings))
        {
            foreach (var kvp in enStrings)
            {
                if (kvp.Key.StartsWith("Tag_") && kvp.Value == displayTag)
                {
                    return kvp.Key.Substring(4) + suffix;
                }
            }
        }

        // If not found, return as-is (might already be internal name)
        return displayTag + suffix;
    }
}

public class Announcer
{
    public string Name { get; set; } = "NewAnnouncer";
    public Color BorderColor { get; set; } = Color.FromArgb(255, 172, 219, 242);
    public bool Random { get; set; } = false;
    public int Quantity { get; set; } = 1;
    public bool Expression { get; set; } = true;
    public bool AutoGenerateTags { get; set; } = true;
    // language -> tag -> text
    public Dictionary<string, Dictionary<string, string>> Texts { get; } = new Dictionary<string, Dictionary<string, string>>();
    // tag -> source image path (on disk)
    public Dictionary<string, string> AssignedImages { get; } = new Dictionary<string, string>();
    // tag -> source sound path (on disk)
    public Dictionary<string, string> AssignedSounds { get; } = new Dictionary<string, string>();
    // announcer image path
    public string AnnouncerImage { get; set; } = "";
    // UI image path
    public string UIImage { get; set; } = "";


    public Announcer Clone()
    {
        var clone = new Announcer
        {
            Name = this.Name,
            BorderColor = this.BorderColor,
            Random = this.Random,
            Quantity = this.Quantity,
            Expression = this.Expression,
            AutoGenerateTags = this.AutoGenerateTags,
            AnnouncerImage = this.AnnouncerImage,
            UIImage = this.UIImage
        };
        foreach (var lang in this.Texts)
        {
            clone.Texts[lang.Key] = new Dictionary<string, string>(lang.Value);
        }
        foreach (var img in this.AssignedImages)
        {
            clone.AssignedImages[img.Key] = img.Value;
        }
        foreach (var snd in this.AssignedSounds)
        {
            clone.AssignedSounds[snd.Key] = snd.Value;
        }
        return clone;
    }
}

public class MainForm : Form
{
    string[] Tags = new[] {
        "Click","StartWork","HalfOverWork","OverWork","AgentDie","AgentPanic",
        "OnAgentPanicReturn","AllDie","OnGetEGOgift","CounterToZero","Suppress",
        "QliphothMeltdown","Rabbit","RabbitReturn","Idle","OnOverWork","AgentHurt"
    };

    // Placeholder arrays for raw values (used when inserting)
    static readonly string[] AgentPlaceholders = new[] { "#0", "[#CurrentSefira_LEB]", "[#AgentLevel_LEB]" };
    static readonly string[] AbnormalityPlaceholders = new[] { "$0", "[$CurrentSefira_LEB]", "[$RiskLevel_LEB]" };
    static readonly string[] GlobalPlaceholders = new[] { "[PlayerHour_LEB]", "[PlayerMinute_LEB]", "[PlayerSecond_LEB]", "[NowEnergy_LEB]", "[NeedEnergy_LEB]", "[StillShort_LEB]" };

    List<Announcer> announcers = new List<Announcer>();

    ListBox lstAnnouncers;
    ListBox lstTags;
    TextBox txtTagText;
    ComboBox cmbLanguage;
    ComboBox cmbBaseTags;
    Button btnAddAnnouncer, btnRemoveAnnouncer, btnSetFolder, btnAssignImage, btnRemoveImage, btnAssignSound, btnRemoveSound, btnPlaySound, btnSaveMod, btnLoadMod, btnAddTag, btnRemoveTag, btnGenerateAllTags, btnAssignAnnouncerImage, btnAssignUIImage, btnAddLanguage, btnRemoveLanguage, btnChangeColor, btnAddPlaceholder;
    ListBox lstAgentPlaceholders, lstAbnormalityPlaceholders, lstGlobalPlaceholders;
    Label lblFolder, lblImageAssigned, lblSoundAssigned, lblAnnouncerImage, lblUIImage, lblTagLimit, lblDeleteHint, lblLanguage, lblToolLanguage, lblQuantity, lblAlpha, lblFontSize, lblInstructions, lblDoubleClickHint;
    PictureBox pbAnnouncerImage, pbUIImage, pbTagImage;
    Panel pnlColorPreview;
    TextBox txtAnnouncerName;
    CheckBox chkRandom, chkAutoGenerateTags, chkTranslatePlaceholders, chkEnableUndoRedo, chkNormalizeAudio;
    NumericUpDown nudQuantity, nudAlpha, nudFontSize;
    TrackBar trkCompression;
    Label lblCompressionValue;
    ComboBox cmbToolLanguage;
    string saveFolder = "";
    WaveOutEvent? waveOut;
    AudioFileReader? audioReader;
    bool isUpdatingAnnouncerSelection = false;
    bool isInitializing = true;

    // Undo/Redo state (one step only)
    List<Announcer>? undoState = null;
    List<Announcer>? redoState = null;
    Button btnUndo, btnRedo;

    static readonly string SettingsFilePath = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "settings.json"
    );

    public MainForm()
    {
        Text = "Announcer Studio";
        Width = 1210; Height = 870;
        StartPosition = FormStartPosition.CenterScreen;

        lstAnnouncers = new ListBox() { Left = 10, Top = 40, Width = 200, Height = 420 };
        lstAnnouncers.SelectedIndexChanged += (s,e)=> RefreshAnnouncerSelection();
        Controls.Add(lstAnnouncers);

        btnAddAnnouncer = new Button() { Left = 10, Top = 470, Width = 95, Text = "Add" };
        btnRemoveAnnouncer = new Button() { Left = 115, Top = 470, Width = 95, Text = "Remove", Enabled = false };
        btnAddAnnouncer.Click += (s,e)=>{ AddAnnouncer(); };
        btnRemoveAnnouncer.Click += (s,e)=>{ RemoveAnnouncer(); };
        Controls.Add(btnAddAnnouncer); Controls.Add(btnRemoveAnnouncer);

        txtAnnouncerName = new TextBox() { Left = 10, Top = 500, Width = 200, Text = "" };
        txtAnnouncerName.TextChanged += (s,e)=> UpdateAnnouncerName();
        Controls.Add(txtAnnouncerName);

        chkRandom = new CheckBox() { Left = 10, Top = 530, Width = 100, Text = "Random" };
        chkRandom.CheckedChanged += (s,e)=> UpdateAnnouncerSettings();
        Controls.Add(chkRandom);

        nudQuantity = new NumericUpDown() { Left = 120, Top = 530, Width = 60, Minimum = 1, Maximum = 10, Value = 1 };
        nudQuantity.ValueChanged += (s,e)=> UpdateAnnouncerSettings();
        Controls.Add(nudQuantity);

        chkAutoGenerateTags = new CheckBox() { Left = 185, Top = 530, Width = 130, Text = "Auto-generate" };
        chkAutoGenerateTags.CheckedChanged += (s,e)=> UpdateAnnouncerSettings();
        Controls.Add(chkAutoGenerateTags);

        btnAssignAnnouncerImage = new Button() { Left = 10, Top = 590, Width = 150, Text = "Assign Announcer.png" };
        btnAssignAnnouncerImage.Click += (s,e)=> AssignAnnouncerImage();
        Controls.Add(btnAssignAnnouncerImage);

        btnAssignUIImage = new Button() { Left = 170, Top = 590, Width = 150, Text = "Assign UI.png" };
        btnAssignUIImage.Click += (s,e)=> AssignUIImage();
        Controls.Add(btnAssignUIImage);

        btnChangeColor = new Button() { Left = 10, Top = 560, Width = 150, Text = "Change Border Color" };
        btnChangeColor.Click += (s,e)=> ChangeBorderColor();
        Controls.Add(btnChangeColor);

        nudAlpha = new NumericUpDown() { Left = 170, Top = 560, Width = 80, Minimum = 0, Maximum = 255, DecimalPlaces = 0, Value = 255 };
        nudAlpha.ValueChanged += (s,e)=> { UpdateAnnouncerSettings(); UpdateColorPreview(); };
        Controls.Add(nudAlpha);

        pnlColorPreview = new Panel() { Left = 260, Top = 560, Width = 60, Height = 21, BorderStyle = BorderStyle.FixedSingle };
        pnlColorPreview.Paint += PnlColorPreview_Paint;
        Controls.Add(pnlColorPreview);

        lblAnnouncerImage = new Label() { Left = 10, Top = 590, Width = 150, Text = "No image assigned" };
        Controls.Add(lblAnnouncerImage);

        lblUIImage = new Label() { Left = 170, Top = 590, Width = 150, Text = "No image assigned" };
        Controls.Add(lblUIImage);

        pbAnnouncerImage = new PictureBox() { Left = 10, Top = 620, Width = 200, Height = 200, BorderStyle = BorderStyle.FixedSingle, SizeMode = PictureBoxSizeMode.Zoom };
        Controls.Add(pbAnnouncerImage);

        pbUIImage = new PictureBox() { Left = 220, Top = 620, Width = 200, Height = 200, BorderStyle = BorderStyle.FixedSingle, SizeMode = PictureBoxSizeMode.Zoom };
        Controls.Add(pbUIImage);

        btnSetFolder = new Button() { Left = 10, Top = 10, Width = 200, Text = "Select Save Folder..." };
        btnSetFolder.Click += (s,e)=> { SelectSaveFolder(); };
        Controls.Add(btnSetFolder);

        lblFolder = new Label() { Left = 220, Top = 15, Width = 740, Text = "No folder selected" };
        Controls.Add(lblFolder);

        lstTags = new ListBox() { Left = 220, Top = 40, Width = 220, Height = 380, DrawMode = DrawMode.OwnerDrawFixed };
        lstTags.SelectedIndexChanged += (s,e)=> LoadSelectedTag();
        lstTags.DrawItem += LstTags_DrawItem;
        lstTags.KeyDown += LstTags_KeyDown;
        Controls.Add(lstTags);

        btnRemoveTag = new Button() { Left = 220, Top = 425, Width = 110, Text = "Remove Tag" };
        btnRemoveTag.Click += (s,e)=> RemoveTag();
        Controls.Add(btnRemoveTag);

        btnGenerateAllTags = new Button() { Left = 335, Top = 425, Width = 105, Text = "Generate All" };
        btnGenerateAllTags.Click += (s,e)=> GenerateAllTags();
        Controls.Add(btnGenerateAllTags);

        cmbBaseTags = new ComboBox() { Left = 220, Top = 455, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbBaseTags.Items.AddRange(Tags);
        cmbBaseTags.SelectedIndexChanged += (s,e)=> UpdateTagLimitLabel();
        Controls.Add(cmbBaseTags);

        btnAddTag = new Button() { Left = 380, Top = 455, Width = 60, Text = "Add" };
        btnAddTag.Click += (s,e)=> AddTag();
        Controls.Add(btnAddTag);

        lblTagLimit = new Label() { Left = 220, Top = 480, Width = 220, Text = "", ForeColor = Color.Red };
        Controls.Add(lblTagLimit);

        lblDeleteHint = new Label() { Left = 220, Top = 500, Width = 220, Text = "Press Delete to remove tag", ForeColor = Color.Gray, Font = new Font(this.Font.FontFamily, 8) };
        Controls.Add(lblDeleteHint);

        lblLanguage = new Label() { Left = 450, Top = 44, Width = 60, Text = "Language:" };
        Controls.Add(lblLanguage);
        cmbLanguage = new ComboBox() { Left = 520, Top = 40, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
        // Start empty - user adds languages via "Add Language" button
        cmbLanguage.SelectedIndexChanged += (s,e)=> OnLanguageChanged();
        Controls.Add(cmbLanguage);

        btnAddLanguage = new Button() { Left = 625, Top = 40, Width = 90, Text = "Add" };
        btnAddLanguage.Click += (s,e)=> AddLanguage();
        Controls.Add(btnAddLanguage);

        btnRemoveLanguage = new Button() { Left = 720, Top = 40, Width = 60, Text = "Delete", Enabled = false };
        btnRemoveLanguage.Click += (s,e)=> RemoveLanguage();
        Controls.Add(btnRemoveLanguage);

        txtTagText = new TextBox() { Left = 450, Top = 80, Width = 510, Height = 240, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font(Font.FontFamily, 18) };
        Controls.Add(txtTagText);

        var lblFontSize = new Label() { Left = 970, Top = 44, Width = 60, Text = "Font Size:" };
        Controls.Add(lblFontSize);

        nudFontSize = new NumericUpDown() { Left = 1035, Top = 40, Width = 60, Minimum = 8, Maximum = 24, Value = 18, DecimalPlaces = 0 };
        nudFontSize.ValueChanged += (s, e) => UpdateTagTextFontSize();
        Controls.Add(nudFontSize);

        // Placeholder lists with add button
        // Agent placeholders: #0, [#CurrentSefira_LEB], [#AgentLevel_LEB]
        lstAgentPlaceholders = new ListBox() { Left = 970, Top = 80, Width = 160, Height = 50 };
        lstAgentPlaceholders.Items.AddRange(new object[] { "#0", "[#CurrentSefira_LEB]", "[#AgentLevel_LEB]" });
        lstAgentPlaceholders.DoubleClick += (s, e) => AddSelectedPlaceholder(lstAgentPlaceholders);
        Controls.Add(lstAgentPlaceholders);

        // Abnormality placeholders: $0, [$CurrentSefira_LEB], [$RiskLevel_LEB]
        lstAbnormalityPlaceholders = new ListBox() { Left = 970, Top = 135, Width = 160, Height = 50 };
        lstAbnormalityPlaceholders.Items.AddRange(new object[] { "$0", "[$CurrentSefira_LEB]", "[$RiskLevel_LEB]" });
        lstAbnormalityPlaceholders.DoubleClick += (s, e) => AddSelectedPlaceholder(lstAbnormalityPlaceholders);
        Controls.Add(lstAbnormalityPlaceholders);

        // Global placeholders (always available)
        lstGlobalPlaceholders = new ListBox() { Left = 970, Top = 190, Width = 160, Height = 100 };
        lstGlobalPlaceholders.Items.AddRange(new object[] { "[PlayerHour_LEB]", "[PlayerMinute_LEB]", "[PlayerSecond_LEB]", "[NowEnergy_LEB]", "[NeedEnergy_LEB]", "[StillShort_LEB]" });
        lstGlobalPlaceholders.DoubleClick += (s, e) => AddSelectedPlaceholder(lstGlobalPlaceholders);
        Controls.Add(lstGlobalPlaceholders);

        btnAddPlaceholder = new Button() { Left = 1135, Top = 150, Width = 50, Text = "Add" };
        btnAddPlaceholder.Click += (s, e) => AddSelectedPlaceholderFromAnyList();
        Controls.Add(btnAddPlaceholder);

        lblDoubleClickHint = new Label() { Left = 1135, Top = 175, Width = 60, Height = 30, Text = "or Double\nClick", ForeColor = Color.Gray, Font = new Font(Font.FontFamily, 7) };
        Controls.Add(lblDoubleClickHint);

        chkTranslatePlaceholders = new CheckBox() { Left = 970, Top = 295, Width = 180, Text = "Translate Placeholders", Checked = true };
        chkTranslatePlaceholders.CheckedChanged += (s, e) => UpdatePlaceholderListItems();
        Controls.Add(chkTranslatePlaceholders);

        btnAssignImage = new Button() { Left = 450, Top = 330, Width = 100, Text = "Assign Image..." };
        btnAssignImage.Click += (s,e)=> AssignImageToSelectedTag();
        Controls.Add(btnAssignImage);

        btnRemoveImage = new Button() { Left = 555, Top = 330, Width = 25, Text = "×", Visible = false };
        btnRemoveImage.Click += (s,e)=> RemoveImageFromSelectedTag();
        Controls.Add(btnRemoveImage);

        lblImageAssigned = new Label() { Left = 585, Top = 335, Width = 195, Text = "No image assigned" };
        Controls.Add(lblImageAssigned);

        btnAssignSound = new Button() { Left = 790, Top = 330, Width = 100, Text = "Assign Sound..." };
        btnAssignSound.Click += (s,e)=> AssignSoundToSelectedTag();
        Controls.Add(btnAssignSound);

        btnRemoveSound = new Button() { Left = 895, Top = 330, Width = 25, Text = "×", Visible = false };
        btnRemoveSound.Click += (s,e)=> RemoveSoundFromSelectedTag();
        Controls.Add(btnRemoveSound);

        btnPlaySound = new Button() { Left = 925, Top = 330, Width = 40, Text = "▶" };
        btnPlaySound.Click += (s,e)=> PlayAssignedSound();
        Controls.Add(btnPlaySound);

        lblSoundAssigned = new Label() { Left = 970, Top = 335, Width = 180, Text = "No sound assigned" };
        Controls.Add(lblSoundAssigned);

        chkNormalizeAudio = new CheckBox() { Left = 790, Top = 355, Width = 120, Text = "Normalize Audio", Checked = true, Visible = false };
        chkNormalizeAudio.CheckedChanged += (s, e) => {
            if (!isInitializing) SaveSettings();
            trkCompression.Enabled = chkNormalizeAudio.Checked;
            lblCompressionValue.Enabled = chkNormalizeAudio.Checked;
        };
        Controls.Add(chkNormalizeAudio);

        trkCompression = new TrackBar() { Left = 915, Top = 352, Width = 120, Minimum = 0, Maximum = 100, Value = 50, TickFrequency = 25, Visible = false };
        trkCompression.ValueChanged += (s, e) => {
            lblCompressionValue.Text = $"{trkCompression.Value}%";
            if (!isInitializing) SaveSettings();
        };
        Controls.Add(trkCompression);

        lblCompressionValue = new Label() { Left = 1040, Top = 358, Width = 40, Text = "50%", Visible = false };
        Controls.Add(lblCompressionValue);

        pbTagImage = new PictureBox() { Left = 450, Top = 360, Width = 200, Height = 200, BorderStyle = BorderStyle.FixedSingle, SizeMode = PictureBoxSizeMode.Zoom };
        Controls.Add(pbTagImage);

        btnSaveMod = new Button() { Left = 450, Top = 570, Width = 120, Text = "Save Mod" };
        btnSaveMod.Click += (s,e)=> SaveMod();
        Controls.Add(btnSaveMod);

        btnLoadMod = new Button() { Left = 580, Top = 570, Width = 120, Text = "Load Existing" };
        btnLoadMod.Click += (s,e)=> LoadExistingMod();
        Controls.Add(btnLoadMod);

        lblToolLanguage = new Label() { Left = 710, Top = 574, Width = 90, Text = "Tool Language:" };
        Controls.Add(lblToolLanguage);

        cmbToolLanguage = new ComboBox() { Left = 805, Top = 570, Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbToolLanguage.SelectedIndexChanged += (s, e) => OnToolLanguageChanged();
        Controls.Add(cmbToolLanguage);

        btnUndo = new Button() { Left = 935, Top = 570, Width = 60, Text = "← Undo", Enabled = false, Visible = false };
        btnUndo.Click += (s, e) => Undo();
        Controls.Add(btnUndo);

        btnRedo = new Button() { Left = 1000, Top = 570, Width = 60, Text = "Redo →", Enabled = false, Visible = false };
        btnRedo.Click += (s, e) => Redo();
        Controls.Add(btnRedo);

        chkEnableUndoRedo = new CheckBox() { Left = 935, Top = 595, Width = 150, Text = "Enable Undo/Redo", Checked = false };
        chkEnableUndoRedo.CheckedChanged += (s, e) => {
            btnUndo.Visible = chkEnableUndoRedo.Checked;
            btnRedo.Visible = chkEnableUndoRedo.Checked;
            SaveSettings();
        };
        Controls.Add(chkEnableUndoRedo);

        lblInstructions = new Label() { Left = 450, Top = 620, Width = 520, Height = 120, Text = "", AutoSize=false };
        Controls.Add(lblInstructions);

        // Load languages from JSON files, then load settings and apply localization
        Localization.LoadLanguages();
        PopulateLanguageComboBox();
        LoadSettings();
        ApplyLocalization();

        // add initial announcer
        AddAnnouncer();

        // Initialization complete
        isInitializing = false;
    }

    void PopulateLanguageComboBox()
    {
        cmbToolLanguage.Items.Clear();
        foreach (var langCode in Localization.GetAvailableLanguages())
        {
            cmbToolLanguage.Items.Add(Localization.GetLanguageName(langCode));
        }
        if (cmbToolLanguage.Items.Count > 0)
        {
            cmbToolLanguage.SelectedIndex = 0;
        }
    }

    void OnToolLanguageChanged()
    {
        if (cmbToolLanguage.SelectedIndex < 0) return;
        var languages = Localization.GetAvailableLanguages();
        if (cmbToolLanguage.SelectedIndex < languages.Count)
        {
            Localization.CurrentLanguage = languages[cmbToolLanguage.SelectedIndex];
            ApplyLocalization();
            // Don't save settings during initialization
            if (!isInitializing)
            {
                SaveSettings();
            }
        }
    }

    void SaveSettings()
    {
        try
        {
            var settings = new AppSettings
            {
                Language = Localization.CurrentLanguage,
                TranslatePlaceholders = chkTranslatePlaceholders.Checked,
                EnableUndoRedo = chkEnableUndoRedo.Checked,
                NormalizeAudio = chkNormalizeAudio.Checked,
                CompressionLevel = trkCompression.Value
            };
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch
        {
            // Silently ignore save errors
        }
    }

    void LoadSettings()
    {
        try
        {
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                if (settings != null)
                {
                    var languages = Localization.GetAvailableLanguages();
                    var idx = languages.IndexOf(settings.Language);
                    if (idx >= 0)
                    {
                        Localization.CurrentLanguage = settings.Language;
                        cmbToolLanguage.SelectedIndex = idx;
                    }
                    chkTranslatePlaceholders.Checked = settings.TranslatePlaceholders;
                    chkEnableUndoRedo.Checked = settings.EnableUndoRedo;
                    btnUndo.Visible = settings.EnableUndoRedo;
                    btnRedo.Visible = settings.EnableUndoRedo;
                    chkNormalizeAudio.Checked = settings.NormalizeAudio;
                    trkCompression.Value = Math.Clamp(settings.CompressionLevel, 0, 100);
                    lblCompressionValue.Text = $"{trkCompression.Value}%";
                    trkCompression.Enabled = settings.NormalizeAudio;
                    lblCompressionValue.Enabled = settings.NormalizeAudio;
                }
            }
        }
        catch
        {
            // Silently ignore load errors, use defaults
        }
    }


    List<Announcer> CloneAnnouncers()
    {
        return announcers.Select(a => a.Clone()).ToList();
    }

    void SaveStateForUndo()
    {
        undoState = CloneAnnouncers();
        redoState = null;
        UpdateUndoRedoButtons();
    }

    void Undo()
    {
        if (undoState == null) return;
        redoState = CloneAnnouncers();
        announcers = undoState;
        undoState = null;
        RefreshAnnouncersListBox();
        UpdateUndoRedoButtons();
    }

    void Redo()
    {
        if (redoState == null) return;
        undoState = CloneAnnouncers();
        announcers = redoState;
        redoState = null;
        RefreshAnnouncersListBox();
        UpdateUndoRedoButtons();
    }

    void UpdateUndoRedoButtons()
    {
        btnUndo.Enabled = undoState != null;
        btnRedo.Enabled = redoState != null;
    }

    void RefreshAnnouncersListBox()
    {
        isUpdatingAnnouncerSelection = true;
        var selectedIdx = lstAnnouncers.SelectedIndex;
        lstAnnouncers.Items.Clear();
        foreach (var a in announcers)
        {
            lstAnnouncers.Items.Add(a.Name);
        }
        if (selectedIdx >= 0 && selectedIdx < lstAnnouncers.Items.Count)
        {
            lstAnnouncers.SelectedIndex = selectedIdx;
        }
        else if (lstAnnouncers.Items.Count > 0)
        {
            lstAnnouncers.SelectedIndex = 0;
        }
        isUpdatingAnnouncerSelection = false;
        RefreshAnnouncerSelection();
    }

    void ApplyLocalization()
    {
        Text = Localization.Get("WindowTitle");
        btnSetFolder.Text = Localization.Get("SelectSaveFolder");
        if (string.IsNullOrEmpty(saveFolder))
            lblFolder.Text = Localization.Get("NoFolderSelected");
        btnAddAnnouncer.Text = Localization.Get("Add");
        btnRemoveAnnouncer.Text = Localization.Get("Remove");
        btnRemoveTag.Text = Localization.Get("RemoveTag");
        btnGenerateAllTags.Text = Localization.Get("GenerateAllTags");
        lblDeleteHint.Text = Localization.Get("DeleteHint");
        lblLanguage.Text = Localization.Get("Language");
        btnAddLanguage.Text = Localization.Get("AddLanguageShort");
        btnRemoveLanguage.Text = Localization.Get("RemoveLanguage");
        btnAddTag.Text = Localization.Get("Add");
        btnAssignImage.Text = Localization.Get("AssignImage");
        btnAssignSound.Text = Localization.Get("AssignSound");
        btnSaveMod.Text = Localization.Get("SaveMod");
        btnLoadMod.Text = Localization.Get("LoadExisting");
        lblInstructions.Text = Localization.Get("Instructions");
        chkRandom.Text = Localization.Get("Random");
        chkAutoGenerateTags.Text = Localization.Get("AutoGenerateTags");
        btnAssignAnnouncerImage.Text = Localization.Get("AssignAnnouncerPng");
        btnAssignUIImage.Text = Localization.Get("AssignUIPng");
        btnChangeColor.Text = Localization.Get("ChangeBorderColor");
        lblToolLanguage.Text = Localization.Get("ToolLanguage");
        btnUndo.Text = "← " + Localization.Get("Undo");
        btnRedo.Text = Localization.Get("Redo") + " →";
        btnAddPlaceholder.Text = Localization.Get("Add");
        chkTranslatePlaceholders.Text = Localization.Get("TranslatePlaceholders");
        chkEnableUndoRedo.Text = Localization.Get("EnableUndoRedo");
        chkNormalizeAudio.Text = Localization.Get("NormalizeAudio");

        // Update placeholder lists if translation is enabled
        UpdatePlaceholderListItems();

        // Update cmbBaseTags with localized tag names
        var selectedBaseTag = cmbBaseTags.SelectedIndex >= 0 ? Tags[cmbBaseTags.SelectedIndex] : null;
        cmbBaseTags.Items.Clear();
        foreach (var baseTag in Tags)
        {
            cmbBaseTags.Items.Add(Localization.GetTagDisplayName(baseTag));
        }
        if (selectedBaseTag != null)
        {
            cmbBaseTags.SelectedIndex = Array.IndexOf(Tags, selectedBaseTag);
        }
        else if (cmbBaseTags.Items.Count > 0)
        {
            cmbBaseTags.SelectedIndex = 0;
        }

        // Refresh tag list to update display names
        lstTags.Invalidate();

        // Update dynamic labels
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx >= 0)
        {
            var a = announcers[aidx];
            if (string.IsNullOrEmpty(a.AnnouncerImage))
                lblAnnouncerImage.Text = Localization.Get("NoImageAssigned");
            if (string.IsNullOrEmpty(a.UIImage))
                lblUIImage.Text = Localization.Get("NoImageAssigned");
        }

        var selectedTag = lstTags.SelectedItem?.ToString();
        if (aidx >= 0 && !string.IsNullOrEmpty(selectedTag))
        {
            var a = announcers[aidx];
            if (!a.AssignedImages.ContainsKey(selectedTag) || string.IsNullOrEmpty(a.AssignedImages[selectedTag]))
                lblImageAssigned.Text = Localization.Get("NoImageAssigned");
            if (!a.AssignedSounds.ContainsKey(selectedTag) || string.IsNullOrEmpty(a.AssignedSounds[selectedTag]))
                lblSoundAssigned.Text = Localization.Get("NoSoundAssigned");
        }
    }

    void AddAnnouncer()
    {
        SaveStateForUndo();
        var a = new Announcer();
        a.Name = "NewAnnouncer" + (announcers.Count + 1);
        // Start with no languages - user must add one via "Add Language" button
        announcers.Add(a);
        lstAnnouncers.Items.Add(a.Name);
        lstAnnouncers.SelectedIndex = lstAnnouncers.Items.Count - 1;
        btnRemoveAnnouncer.Enabled = announcers.Count > 1;
    }

    void OnLanguageChanged()
    {
        if (isUpdatingAnnouncerSelection) return;
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) return;
        
        var a = announcers[idx];
        var lang = cmbLanguage.SelectedItem?.ToString();
        
        // Update remove language button state
        btnRemoveLanguage.Enabled = cmbLanguage.Items.Count > 0;
        
        if (string.IsNullOrEmpty(lang))
        {
            lstTags.Items.Clear();
            LoadSelectedTag();
            return;
        }
        
        // Only show tags if the language exists for this announcer
        lstTags.Items.Clear();
        if (a.Texts.TryGetValue(lang, out var texts))
        {
            // Sort tags by base tag order, then by number
            var sortedTags = texts.Keys
                .OrderBy(t => {
                    string bt;
                    if (t.Contains("_") && int.TryParse(t.Split('_').Last(), out _))
                    {
                        var parts = t.Split('_');
                        bt = string.Join("_", parts.Take(parts.Length - 1));
                    }
                    else
                    {
                        bt = t;
                    }
                    int index = Array.IndexOf(Tags, bt);
                    return index >= 0 ? index : int.MaxValue;
                })
                .ThenBy(t => {
                    if (t.Contains("_") && int.TryParse(t.Split('_').Last(), out var n))
                        return n;
                    return 0;
                });
            
            foreach (var tag in sortedTags)
            {
                lstTags.Items.Add(tag);
            }
        }
        LoadSelectedTag();
    }

    void UpdateAnnouncerName()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0 || isUpdatingAnnouncerSelection) return;
        var a = announcers[idx];
        a.Name = txtAnnouncerName.Text;
        lstAnnouncers.Items[idx] = a.Name;
    }

    void UpdateAnnouncerSettings()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0 || isUpdatingAnnouncerSelection) return;
        var a = announcers[idx];
        var oldQuantity = a.Quantity;
        var oldRandom = a.Random;
        var oldAutoGenerateTags = a.AutoGenerateTags;
        var oldAlpha = a.BorderColor.A;
        
        bool hasChanges = chkRandom.Checked != oldRandom 
            || (int)nudQuantity.Value != oldQuantity 
            || chkAutoGenerateTags.Checked != oldAutoGenerateTags
            || (int)nudAlpha.Value != oldAlpha;
        
        if (hasChanges)
        {
            SaveStateForUndo();
        }
        
        a.Random = chkRandom.Checked;
        a.Quantity = (int)nudQuantity.Value;
        a.AutoGenerateTags = chkAutoGenerateTags.Checked;
        if (a.Quantity != oldQuantity || a.Random != oldRandom || a.AutoGenerateTags != oldAutoGenerateTags)
        {
            RenumberTags(a);

            // If AutoGenerateTags is now enabled, ensure all tags exist
            if (a.AutoGenerateTags)
            {
                var lang = cmbLanguage.Text.Trim();
                if (!string.IsNullOrEmpty(lang))
                {
                    EnsureAllTagsExist(a, lang);
                    RefreshTagList(a, lang);
                }
            }
        }
        a.BorderColor = Color.FromArgb((int)nudAlpha.Value, a.BorderColor.R, a.BorderColor.G, a.BorderColor.B);
        UpdateTagControlsVisibility();
    }

    void UpdateTagControlsVisibility()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) return;
        var a = announcers[idx];

        // Hide tag dialog, Add button, Remove button, and Delete hint when Auto-Generate is enabled
        bool visible = !a.AutoGenerateTags;
        cmbBaseTags.Visible = visible;
        btnAddTag.Visible = visible;
        btnRemoveTag.Visible = visible;
        lblDeleteHint.Visible = visible;
    }

    void RenumberTags(Announcer a)
    {
        var currentLang = cmbLanguage.Text.Trim();
        if (string.IsNullOrEmpty(currentLang)) return;
        if (!a.Texts.ContainsKey(currentLang)) return;

        var images = a.AssignedImages;
        var sounds = a.AssignedSounds;

        // Process all languages
        foreach (var langKv in a.Texts)
        {
            var texts = langKv.Value;

            // Collect all base tags and their variants
            var baseTags = new Dictionary<string, List<string>>();
            foreach (var tag in texts.Keys.ToArray())
            {
                string baseTag = GetBaseTag(tag);
                if (!baseTags.ContainsKey(baseTag)) baseTags[baseTag] = new List<string>();
                baseTags[baseTag].Add(tag);
            }

            foreach (var kv in baseTags)
            {
                var baseTag = kv.Key;
                var variants = kv.Value;

                if (!a.Random)
                {
                    // Convert to non-random: keep only baseTag
                    // Save existing data (prefer _1, then baseTag)
                    string? savedText = null;
                    if (texts.TryGetValue(baseTag + "_1", out var t1)) savedText = t1;
                    else if (texts.TryGetValue(baseTag, out var t0)) savedText = t0;

                    // Remove all numbered variants and keep/create baseTag
                    foreach (var v in variants.ToList())
                    {
                        if (v != baseTag)
                        {
                            texts.Remove(v);
                        }
                    }
                    if (savedText != null)
                    {
                        texts[baseTag] = savedText;
                    }
                    else if (!texts.ContainsKey(baseTag))
                    {
                        texts[baseTag] = "";
                    }
                }
                else
                {
                    // Convert to random: use numbered tags (_1, _2, etc.)
                    // Save existing data
                    var existingTexts = new Dictionary<int, string>();
                    foreach (var v in variants)
                    {
                        if (v == baseTag)
                        {
                            existingTexts[1] = texts[v];
                        }
                        else if (v.StartsWith(baseTag + "_") && int.TryParse(v.Substring(baseTag.Length + 1), out var n))
                        {
                            existingTexts[n] = texts[v];
                        }
                    }

                    // Remove all old variants
                    foreach (var v in variants.ToList())
                    {
                        texts.Remove(v);
                    }

                    // Add numbered tags based on AutoGenerateTags setting
                    if (a.AutoGenerateTags)
                    {
                        for (int i = 1; i <= a.Quantity; i++)
                        {
                            var tag = baseTag + "_" + i;
                            texts[tag] = existingTexts.TryGetValue(i, out var t) ? t : "";
                        }
                    }
                    else
                    {
                        // Manual mode: only create tags for existing data
                        foreach (var kvText in existingTexts)
                        {
                            if (kvText.Key <= a.Quantity)
                            {
                                texts[baseTag + "_" + kvText.Key] = kvText.Value;
                            }
                        }
                    }
                }
            }
        }

        // Process images (shared across languages)
        ProcessAssignmentsForRandom(a, images);
        // Process sounds (shared across languages)
        ProcessAssignmentsForRandom(a, sounds);

        // Update lstTags for current language
        RefreshTagList(a, currentLang);
        UpdateTagLimitLabel();
    }

    void ProcessAssignmentsForRandom(Announcer a, Dictionary<string, string> assignments)
    {
        // Collect all base tags from assignments
        var baseTags = new Dictionary<string, List<string>>();
        foreach (var tag in assignments.Keys.ToArray())
        {
            string baseTag = GetBaseTag(tag);
            if (!baseTags.ContainsKey(baseTag)) baseTags[baseTag] = new List<string>();
            baseTags[baseTag].Add(tag);
        }

        foreach (var kv in baseTags)
        {
            var baseTag = kv.Key;
            var variants = kv.Value;

            if (!a.Random)
            {
                // Convert to non-random: keep only baseTag
                string? savedValue = null;
                if (assignments.TryGetValue(baseTag + "_1", out var v1)) savedValue = v1;
                else if (assignments.TryGetValue(baseTag, out var v0)) savedValue = v0;

                foreach (var v in variants.ToList())
                {
                    if (v != baseTag) assignments.Remove(v);
                }
                if (savedValue != null)
                {
                    assignments[baseTag] = savedValue;
                }
            }
            else
            {
                // Convert to random: use numbered tags
                var existingValues = new Dictionary<int, string>();
                foreach (var v in variants)
                {
                    if (v == baseTag)
                    {
                        existingValues[1] = assignments[v];
                    }
                    else if (v.StartsWith(baseTag + "_") && int.TryParse(v.Substring(baseTag.Length + 1), out var n))
                    {
                        existingValues[n] = assignments[v];
                    }
                }

                // Remove all old variants
                foreach (var v in variants.ToList())
                {
                    assignments.Remove(v);
                }

                // Restore with numbered tags
                foreach (var kvVal in existingValues)
                {
                    if (kvVal.Key <= a.Quantity)
                    {
                        assignments[baseTag + "_" + kvVal.Key] = kvVal.Value;
                    }
                }
            }
        }
    }

    string GetBaseTag(string tag)
    {
        if (tag.Contains("_") && int.TryParse(tag.Split('_').Last(), out _))
        {
            var parts = tag.Split('_');
            return string.Join("_", parts.Take(parts.Length - 1));
        }
        return tag;
    }

    /// <summary>
    /// Converts internal tag name to XML tag name format.
    /// Internal: "AgentDie" or "AgentDie_1"
    /// XML: "AgentDie_TEXT" or "AgentDie_TEXT_1"
    /// </summary>
    string ConvertToXmlTagName(string internalTag)
    {
        if (internalTag.Contains("_") && int.TryParse(internalTag.Split('_').Last(), out int number))
        {
            // Has number suffix like "AgentDie_1" -> "AgentDie_TEXT_1"
            var baseTag = GetBaseTag(internalTag);
            return $"{baseTag}_TEXT_{number}";
        }
        // No number suffix like "AgentDie" -> "AgentDie_TEXT"
        return $"{internalTag}_TEXT";
    }

    void RefreshTagList(Announcer a, string lang)
    {
        if (!a.Texts.TryGetValue(lang, out var texts)) return;

        lstTags.Items.Clear();
        var allTags = texts.Keys
            .OrderBy(t => {
                string baseTag = GetBaseTag(t);
                int index = Array.IndexOf(Tags, baseTag);
                return index >= 0 ? index : int.MaxValue;
            })
            .ThenBy(t => {
                if (t.Contains("_") && int.TryParse(t.Split('_').Last(), out var n))
                    return n;
                return 0;
            });
        foreach (var tag in allTags)
        {
            lstTags.Items.Add(tag);
        }
    }

    void UpdateTagLimitLabel()
    {
        var selectedIdx = cmbBaseTags.SelectedIndex;
        if (selectedIdx < 0)
        {
            lblTagLimit.Text = "";
            btnAddTag.Enabled = true;
            return;
        }
        var baseTag = Tags[selectedIdx]; // Use internal tag name

        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0)
        {
            lblTagLimit.Text = "";
            btnAddTag.Enabled = true;
            return;
        }

        var a = announcers[aidx];
        if (!a.Random)
        {
            lblTagLimit.Text = "";
            btnAddTag.Enabled = true;
            return;
        }

        // Count existing tags of this type
        int existingCount = 0;
        foreach(var existingTag in lstTags.Items.Cast<string>())
        {
            if (existingTag == baseTag || existingTag.StartsWith(baseTag + "_"))
            {
                existingCount++;
            }
        }

        if (existingCount >= a.Quantity)
        {
            lblTagLimit.Text = string.Format(Localization.Get("MaximumReached"), existingCount, a.Quantity);
            btnAddTag.Enabled = false;
        }
        else
        {
            lblTagLimit.Text = "";
            btnAddTag.Enabled = true;
        }
    }

    void AddTag()
    {
        // Get the internal tag name from the selected index
        var selectedIdx = cmbBaseTags.SelectedIndex;
        if (selectedIdx < 0) return;
        var baseTag = Tags[selectedIdx]; // Use internal tag name from Tags array
        
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) return;
        var a = announcers[aidx];
        var lang = cmbLanguage.Text.Trim();
        if (string.IsNullOrEmpty(lang)) return;
        if (!a.Texts.ContainsKey(lang)) a.Texts[lang] = new Dictionary<string, string>();

        // Count existing tags of this type
        int existingCount = 0;
        foreach(var existingTag in lstTags.Items.Cast<string>())
        {
            if (existingTag == baseTag || existingTag.StartsWith(baseTag + "_"))
            {
                existingCount++;
            }
        }

        // Check if we've reached the Quantity limit
        if (a.Random && existingCount >= a.Quantity)
        {
            MessageBox.Show(string.Format(Localization.Get("LimitReached"), Localization.GetTagDisplayName(baseTag), a.Quantity), Localization.Get("LimitReachedTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SaveStateForUndo();

        // find the next number
        int maxNum = 0;
        foreach(var existingTag in lstTags.Items.Cast<string>())
        {
            if (existingTag.StartsWith(baseTag + "_"))
            {
                var suffix = existingTag.Substring(baseTag.Length + 1);
                if (int.TryParse(suffix, out var num)) maxNum = Math.Max(maxNum, num);
            }
            else if (existingTag == baseTag)
            {
                maxNum = Math.Max(maxNum, 0);
            }
        }
        string newTag;
        // If Random is enabled, always use numbered tags
        if (a.Random)
        {
            newTag = baseTag + "_" + (maxNum + 1);
        }
        else if (maxNum == 0 && !lstTags.Items.Contains(baseTag))
        {
            newTag = baseTag;
        }
        else
        {
            newTag = baseTag + "_" + (maxNum + 1);
        }

        // Add to texts dictionary
        if (!a.Texts[lang].ContainsKey(newTag))
        {
            a.Texts[lang][newTag] = "";
        }

        if (!lstTags.Items.Contains(newTag))
        {
            lstTags.Items.Add(newTag);
        }
        lstTags.SelectedItem = newTag;
        UpdateTagLimitLabel();
    }


    void GenerateAllTags()
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) return;
        
        var a = announcers[aidx];
        var lang = cmbLanguage.Text.Trim();
        if (string.IsNullOrEmpty(lang))
        {
            MessageBox.Show(Localization.Get("NoLanguageSelected"), Localization.Get("Error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        
        SaveStateForUndo();
        
        // Create or get the texts dictionary for this language
        if (!a.Texts.ContainsKey(lang))
            a.Texts[lang] = new Dictionary<string, string>();
        
        var texts = a.Texts[lang];
        
        // Generate all tags based on Random and Quantity settings
        foreach (var baseTag in Tags)
        {
            if (a.Random && a.Quantity > 1)
            {
                // Generate numbered variants
                for (int i = 1; i <= a.Quantity; i++)
                {
                    var tag = baseTag + "_" + i;
                    if (!texts.ContainsKey(tag))
                        texts[tag] = "";
                }
            }
            else
            {
                // Generate single tag
                if (!texts.ContainsKey(baseTag))
                    texts[baseTag] = "";
            }
        }
        
        // Refresh the tag list
        lstTags.Items.Clear();
        var allTags = texts.Keys
            .OrderBy(t => {
                string bt;
                if (t.Contains("_") && int.TryParse(t.Split('_').Last(), out _))
                {
                    var parts = t.Split('_');
                    bt = string.Join("_", parts.Take(parts.Length - 1));
                }
                else
                {
                    bt = t;
                }
                int index = Array.IndexOf(Tags, bt);
                return index >= 0 ? index : int.MaxValue;
            })
            .ThenBy(t => {
                if (t.Contains("_") && int.TryParse(t.Split('_').Last(), out var n))
                {
                    return n;
                }
                return 0;
            });
        
        foreach (var tag in allTags)
        {
            lstTags.Items.Add(tag);
        }

        UpdateTagLimitLabel();
    }

    // Ensures all tags exist for the current language without showing messages or saving undo state
    void EnsureAllTagsExist(Announcer a, string lang)
    {
        if (string.IsNullOrEmpty(lang)) return;
        if (!a.Texts.ContainsKey(lang))
            a.Texts[lang] = new Dictionary<string, string>();

        var texts = a.Texts[lang];

        // Generate all missing tags based on Random and Quantity settings
        foreach (var baseTag in Tags)
        {
            if (a.Random)
            {
                // Generate numbered variants
                for (int i = 1; i <= a.Quantity; i++)
                {
                    var tag = baseTag + "_" + i;
                    if (!texts.ContainsKey(tag))
                        texts[tag] = "";
                }
            }
            else
            {
                // Generate single tag
                if (!texts.ContainsKey(baseTag))
                    texts[baseTag] = "";
            }
        }
    }

    void RemoveAnnouncer()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0 || announcers.Count <= 1) return;
        var result = MessageBox.Show("Are you sure you want to remove this announcer?", "Confirm", MessageBoxButtons.YesNo);
        if (result == DialogResult.Yes)
        {
            SaveStateForUndo();
            announcers.RemoveAt(idx);
            lstAnnouncers.Items.RemoveAt(idx);
            if (lstAnnouncers.Items.Count>0) lstAnnouncers.SelectedIndex = 0;
            btnRemoveAnnouncer.Enabled = announcers.Count > 1;
        }
    }
    
    void RefreshAnnouncerSelection()
    {
        if (isUpdatingAnnouncerSelection) return;
        isUpdatingAnnouncerSelection = true;
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) { isUpdatingAnnouncerSelection = false; return; }
        var a = announcers[idx];
        // reflect name
        lstAnnouncers.Items[idx] = a.Name;
        txtAnnouncerName.Text = a.Name;
        chkRandom.Checked = a.Random;
        nudQuantity.Value = a.Quantity;
        chkAutoGenerateTags.Checked = a.AutoGenerateTags;
        nudAlpha.Value = a.BorderColor.A;
        lblAnnouncerImage.Text = string.IsNullOrEmpty(a.AnnouncerImage) ? "No image assigned" : Path.GetFileName(a.AnnouncerImage);
        lblUIImage.Text = string.IsNullOrEmpty(a.UIImage) ? "No image assigned" : Path.GetFileName(a.UIImage);
        pbAnnouncerImage.Image = string.IsNullOrEmpty(a.AnnouncerImage) ? null : Image.FromFile(a.AnnouncerImage);
        pbUIImage.Image = string.IsNullOrEmpty(a.UIImage) ? null : Image.FromFile(a.UIImage);
        
        // Update language combo box with this announcer's languages
        cmbLanguage.Items.Clear();
        foreach (var lang in a.Texts.Keys)
        {
            cmbLanguage.Items.Add(lang);
        }
        
        // update tag list using current language
        lstTags.Items.Clear();
        var currentLang = cmbLanguage.Text.Trim();
        if (!string.IsNullOrEmpty(currentLang) && a.Texts.TryGetValue(currentLang, out var texts))
        {
            foreach(var tag in texts.Keys)
            {
                lstTags.Items.Add(tag);
            }
        }
        else if (a.Texts.Count > 0)
        {
            // If current language is not found but there are languages, use the first one
            var firstLang = a.Texts.Keys.First();
            cmbLanguage.Text = firstLang;
            if (a.Texts.TryGetValue(firstLang, out var firstTexts))
            {
                foreach(var tag in firstTexts.Keys)
                {
                    lstTags.Items.Add(tag);
                }
            }
        }
        // If no languages exist, lstTags stays empty
        
        LoadSelectedTag();
        isUpdatingAnnouncerSelection = false;
        btnRemoveAnnouncer.Enabled = announcers.Count > 1;
        UpdateTagLimitLabel();
        UpdateTagControlsVisibility();
        UpdateColorPreview();
    }

    void LoadSelectedTag()
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0)
        {
            HideTagEditUI();
            return;
        }
        var a = announcers[aidx];
        var lang = cmbLanguage.Text.Trim();
        if (string.IsNullOrEmpty(lang))
        {
            HideTagEditUI();
            return;
        }
        if (!a.Texts.ContainsKey(lang)) a.Texts[lang] = new Dictionary<string, string>();
        var tag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(tag))
        {
            HideTagEditUI();
            return;
        }

        // Show tag edit UI
        ShowTagEditUI();

        txtTagText.Text = a.Texts[lang].TryGetValue(tag, out var v) ? v : "";

        // Update image assignment display
        bool hasImage = a.AssignedImages.TryGetValue(tag, out var imgPath) && !string.IsNullOrEmpty(imgPath);
        lblImageAssigned.Text = hasImage ? Path.GetFileName(imgPath) : Localization.Get("NoImageAssigned");
        pbTagImage.Image = hasImage ? Image.FromFile(imgPath) : null;
        btnRemoveImage.Visible = hasImage;

        // Update sound assignment display
        bool hasSound = a.AssignedSounds.TryGetValue(tag, out var sndPath) && !string.IsNullOrEmpty(sndPath);
        lblSoundAssigned.Text = hasSound ? Path.GetFileName(sndPath) : Localization.Get("NoSoundAssigned");
        btnRemoveSound.Visible = hasSound;

        // Update placeholder buttons based on selected tag
        UpdatePlaceholderButtons(tag);

        // attach save on text change
        txtTagText.Leave -= TxtTagText_Leave;
        txtTagText.Leave += TxtTagText_Leave;
    }

    void UpdatePlaceholderButtons(string tag)
    {
        // Extract base tag name (remove _1, _2, etc. suffix)
        var baseTag = tag;
        var lastUnderscore = tag.LastIndexOf('_');
        if (lastUnderscore > 0 && int.TryParse(tag.Substring(lastUnderscore + 1), out _))
        {
            baseTag = tag.Substring(0, lastUnderscore);
        }

        // Agent placeholders - only for: AgentDie, AgentPanic, OnAgentPanicReturn, OnGetEGOgift, AgentHurt
        var agentTags = new[] { "AgentDie", "AgentPanic", "OnAgentPanicReturn", "OnGetEGOgift", "AgentHurt" };
        lstAgentPlaceholders.Enabled = agentTags.Contains(baseTag);

        // Abnormality placeholders - only for: CounterToZero, Suppress
        var abnormalityTags = new[] { "CounterToZero", "Suppress" };
        lstAbnormalityPlaceholders.Enabled = abnormalityTags.Contains(baseTag);

        // Global placeholders - always enabled
        lstGlobalPlaceholders.Enabled = true;
    }

    void ShowTagEditUI()
    {
        // Note: lblLanguage, cmbLanguage, btnAddLanguage are always visible
        txtTagText.Visible = true;
        btnAssignImage.Visible = true;
        lblImageAssigned.Visible = true;
        pbTagImage.Visible = true;
        btnAssignSound.Visible = true;
        btnPlaySound.Visible = true;
        lblSoundAssigned.Visible = true;
        chkNormalizeAudio.Visible = true;
        trkCompression.Visible = true;
        lblCompressionValue.Visible = true;
        lstAgentPlaceholders.Visible = true;
        lstAbnormalityPlaceholders.Visible = true;
        lstGlobalPlaceholders.Visible = true;
        btnAddPlaceholder.Visible = true;
        lblDoubleClickHint.Visible = true;
        chkTranslatePlaceholders.Visible = true;
    }

    void HideTagEditUI()
    {
        // Note: lblLanguage, cmbLanguage, btnAddLanguage should always remain visible
        // They are needed to add/select languages even when no tag is selected
        txtTagText.Visible = false;
        btnAssignImage.Visible = false;
        btnRemoveImage.Visible = false;
        lblImageAssigned.Visible = false;
        pbTagImage.Visible = false;
        btnAssignSound.Visible = false;
        btnRemoveSound.Visible = false;
        btnPlaySound.Visible = false;
        lblSoundAssigned.Visible = false;
        chkNormalizeAudio.Visible = false;
        trkCompression.Visible = false;
        lblCompressionValue.Visible = false;
        lstAgentPlaceholders.Visible = false;
        lstAbnormalityPlaceholders.Visible = false;
        lstGlobalPlaceholders.Visible = false;
        btnAddPlaceholder.Visible = false;
        lblDoubleClickHint.Visible = false;
        chkTranslatePlaceholders.Visible = false;
    }

    private void TxtTagText_Leave(object? sender, EventArgs e)
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) return;
        var a = announcers[aidx];
        var lang = cmbLanguage.Text.Trim();
        if (string.IsNullOrEmpty(lang)) return;
        var tag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(tag)) return;
        if (!a.Texts.ContainsKey(lang)) return;

        var oldText = a.Texts[lang].TryGetValue(tag, out var v) ? v : "";
        var newText = txtTagText.Text;
        if (oldText != newText)
        {
            SaveStateForUndo();
            a.Texts[lang][tag] = newText;
        }
    }

    void AddPlaceholder(string placeholder)
    {
        var pos = txtTagText.SelectionStart;
        txtTagText.Text = txtTagText.Text.Insert(pos, placeholder);
        txtTagText.SelectionStart = pos + placeholder.Length;
        txtTagText.Focus();
    }

    void AddSelectedPlaceholder(ListBox listBox)
    {
        if (listBox.SelectedIndex >= 0)
        {
            // Get the raw placeholder value from the corresponding array
            string placeholder;
            if (listBox == lstAgentPlaceholders)
                placeholder = AgentPlaceholders[listBox.SelectedIndex];
            else if (listBox == lstAbnormalityPlaceholders)
                placeholder = AbnormalityPlaceholders[listBox.SelectedIndex];
            else
                placeholder = GlobalPlaceholders[listBox.SelectedIndex];
            AddPlaceholder(placeholder);
        }
    }

    void AddSelectedPlaceholderFromAnyList()
    {
        // Check each list for selection and add if found (use index to get raw value)
        if (lstAgentPlaceholders.Enabled && lstAgentPlaceholders.SelectedIndex >= 0)
        {
            AddPlaceholder(AgentPlaceholders[lstAgentPlaceholders.SelectedIndex]);
        }
        else if (lstAbnormalityPlaceholders.Enabled && lstAbnormalityPlaceholders.SelectedIndex >= 0)
        {
            AddPlaceholder(AbnormalityPlaceholders[lstAbnormalityPlaceholders.SelectedIndex]);
        }
        else if (lstGlobalPlaceholders.SelectedIndex >= 0)
        {
            AddPlaceholder(GlobalPlaceholders[lstGlobalPlaceholders.SelectedIndex]);
        }
    }

    void UpdatePlaceholderListItems()
    {
        bool translate = chkTranslatePlaceholders.Checked;

        // Save current selections
        int agentIdx = lstAgentPlaceholders.SelectedIndex;
        int abnormalityIdx = lstAbnormalityPlaceholders.SelectedIndex;
        int globalIdx = lstGlobalPlaceholders.SelectedIndex;

        // Update Agent placeholders
        lstAgentPlaceholders.Items.Clear();
        for (int i = 0; i < AgentPlaceholders.Length; i++)
        {
            lstAgentPlaceholders.Items.Add(translate ? Localization.Get("Placeholder_Agent_" + i) : AgentPlaceholders[i]);
        }

        // Update Abnormality placeholders
        lstAbnormalityPlaceholders.Items.Clear();
        for (int i = 0; i < AbnormalityPlaceholders.Length; i++)
        {
            lstAbnormalityPlaceholders.Items.Add(translate ? Localization.Get("Placeholder_Abnormality_" + i) : AbnormalityPlaceholders[i]);
        }

        // Update Global placeholders
        lstGlobalPlaceholders.Items.Clear();
        for (int i = 0; i < GlobalPlaceholders.Length; i++)
        {
            lstGlobalPlaceholders.Items.Add(translate ? Localization.Get("Placeholder_Global_" + i) : GlobalPlaceholders[i]);
        }

        // Restore selections
        if (agentIdx >= 0 && agentIdx < lstAgentPlaceholders.Items.Count)
            lstAgentPlaceholders.SelectedIndex = agentIdx;
        if (abnormalityIdx >= 0 && abnormalityIdx < lstAbnormalityPlaceholders.Items.Count)
            lstAbnormalityPlaceholders.SelectedIndex = abnormalityIdx;
        if (globalIdx >= 0 && globalIdx < lstGlobalPlaceholders.Items.Count)
            lstGlobalPlaceholders.SelectedIndex = globalIdx;
    }

    void AddLanguage()
    {
        string lang = Interaction.InputBox(Localization.Get("EnterLanguageCode"), Localization.Get("AddLanguageTitle"));
        if (string.IsNullOrEmpty(lang)) return;
        
        lang = lang.Trim();
        if (string.IsNullOrEmpty(lang)) return;
        
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) return;
        
        var a = announcers[idx];
        
        // Check if this language already exists for this announcer
        if (a.Texts.ContainsKey(lang))
        {
            // Language already exists, just switch to it
            var existingIdx = cmbLanguage.Items.IndexOf(lang);
            if (existingIdx >= 0)
                cmbLanguage.SelectedIndex = existingIdx;
            return;
        }
        
        // New language - save state for undo
        SaveStateForUndo();
        
        // Create tags for this new language
        if (a.AutoGenerateTags)
        {
            var newTexts = new Dictionary<string, string>();
            foreach (var baseTag in Tags)
            {
                if (a.Random && a.Quantity > 1)
                {
                    for (int i = 1; i <= a.Quantity; i++)
                    {
                        newTexts[baseTag + "_" + i] = "";
                    }
                }
                else
                {
                    newTexts[baseTag] = "";
                }
            }
            a.Texts[lang] = newTexts;
        }
        else
        {
            a.Texts[lang] = Tags.ToDictionary(t => t, t => "");
        }
        
        // Add to combo box and select it
        cmbLanguage.Items.Add(lang);
        cmbLanguage.SelectedIndex = cmbLanguage.Items.Count - 1;
        
        // Enable remove button now that we have a language
        btnRemoveLanguage.Enabled = true;
    }


    void RemoveLanguage()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) return;
        
        var lang = cmbLanguage.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(lang)) return;
        
        var a = announcers[idx];
        
        // Confirm deletion
        var result = MessageBox.Show(
            string.Format(Localization.Get("ConfirmDeleteLanguage"), lang),
            Localization.Get("ConfirmDeleteLanguageTitle"),
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);
        
        if (result != DialogResult.Yes) return;
        
        SaveStateForUndo();
        
        // Remove the language from this announcer
        a.Texts.Remove(lang);
        
        // Update combo box
        cmbLanguage.Items.Remove(lang);
        
        // Select another language if available
        if (cmbLanguage.Items.Count > 0)
            cmbLanguage.SelectedIndex = 0;
        else
            lstTags.Items.Clear();
        
        // Update remove button state
        btnRemoveLanguage.Enabled = cmbLanguage.Items.Count > 0;
        
        LoadSelectedTag();
    }

    void UpdateTagTextFontSize()
    {
        var size = (float)nudFontSize.Value;
        txtTagText.Font = new Font(txtTagText.Font.FontFamily, size);
    }

    void ChangeBorderColor()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) return;
        var a = announcers[idx];
        using (var cd = new ColorDialog() { Color = Color.FromArgb(255, a.BorderColor.R, a.BorderColor.G, a.BorderColor.B) })
        {
            if (cd.ShowDialog() == DialogResult.OK)
            {
                SaveStateForUndo();
                a.BorderColor = Color.FromArgb(a.BorderColor.A, cd.Color.R, cd.Color.G, cd.Color.B);
                UpdateColorPreview();
            }
        }
    }


    void UpdateColorPreview()
    {
        pnlColorPreview.Invalidate();
    }

    void PnlColorPreview_Paint(object sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        var rect = pnlColorPreview.ClientRectangle;

        // Draw checkerboard background to show transparency
        int checkSize = 8;
        for (int y = 0; y < rect.Height; y += checkSize)
        {
            for (int x = 0; x < rect.Width; x += checkSize)
            {
                var isLight = ((x / checkSize) + (y / checkSize)) % 2 == 0;
                using (var brush = new SolidBrush(isLight ? Color.White : Color.LightGray))
                {
                    g.FillRectangle(brush, x, y, checkSize, checkSize);
                }
            }
        }

        // Get current color from selected announcer
        var idx = lstAnnouncers.SelectedIndex;
        if (idx >= 0)
        {
            var a = announcers[idx];
            using (var brush = new SolidBrush(a.BorderColor))
            {
                g.FillRectangle(brush, rect);
            }
        }
    }

    void LstTags_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) return;
        var internalTag = lstTags.Items[e.Index].ToString();
        var displayTag = Localization.GetTagDisplayName(internalTag ?? "");
        var idx = lstAnnouncers.SelectedIndex;
        if (idx >= 0)
        {
            var a = announcers[idx];
            var lang = cmbLanguage.Text.Trim();
            var text = !string.IsNullOrEmpty(lang) && a.Texts.TryGetValue(lang, out var texts) && texts.TryGetValue(internalTag ?? "", out var t) ? t : "";
            e.DrawBackground();

            // Use white text when selected, otherwise gray for empty tags and black for non-empty
            Brush brush;
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            if (isSelected)
            {
                brush = Brushes.White;
            }
            else
            {
                brush = string.IsNullOrEmpty(text.Trim()) ? Brushes.Gray : Brushes.Black;
            }

            e.Graphics.DrawString(displayTag, e.Font, brush, e.Bounds);
            e.DrawFocusRectangle();
        }
    }

    void LstTags_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Delete)
        {
            RemoveSelectedTag();
            e.Handled = true;
        }
    }

    void RemoveTag()
    {
        RemoveSelectedTag();
    }

    void RemoveSelectedTag()
    {
        // Only allow deletion from the main tag list (lstTags), not from cmbBaseTags
        var selectedTag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(selectedTag))
        {
            MessageBox.Show(Localization.Get("NoTagSelected"), Localization.Get("NoTagSelectedTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) return;

        var a = announcers[aidx];

        // Cannot remove tags when Auto-Generate is enabled
        if (a.AutoGenerateTags)
        {
            MessageBox.Show(Localization.Get("AutoGenerateEnabled"), Localization.Get("AutoGenerateEnabledTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        SaveStateForUndo();

        var lang = cmbLanguage.Text.Trim();
        if (string.IsNullOrEmpty(lang)) return;

        if (a.Texts.TryGetValue(lang, out var texts))
        {
            texts.Remove(selectedTag);
        }
        a.AssignedImages.Remove(selectedTag);
        lstTags.Items.Remove(selectedTag);
        UpdateTagLimitLabel();
    }

    void AssignImageToSelectedTag()
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) { MessageBox.Show(Localization.Get("AnnouncerOrTagNotSelected")); return; }
        var tag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(tag)) { MessageBox.Show(Localization.Get("TagNotSelected")); return; }
        using var ofd = new OpenFileDialog();
        ofd.Filter = "PNG Files|*.png";
        if (ofd.ShowDialog() != DialogResult.OK) return;

        // Validate and optionally resize image (required: 512x512)
        var imagePath = ShowResizeDialog(ofd.FileName, 512, 512, Localization.Get("PortraitImage"));
        if (imagePath == null) return;

        SaveStateForUndo();
        var a = announcers[aidx];
        a.AssignedImages[tag] = imagePath;
        lblImageAssigned.Text = Path.GetFileName(imagePath);
        pbTagImage.Image = Image.FromFile(imagePath);
        btnRemoveImage.Visible = true;
    }

    void AssignSoundToSelectedTag()
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) { MessageBox.Show(Localization.Get("AnnouncerOrTagNotSelected")); return; }
        var tag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(tag)) { MessageBox.Show(Localization.Get("TagNotSelected")); return; }
        using var ofd = new OpenFileDialog();
        ofd.Filter = "WAV Files|*.wav";
        if (ofd.ShowDialog() != DialogResult.OK) return;

        SaveStateForUndo();
        var a = announcers[aidx];
        var sourceFile = ofd.FileName;

        // Auto-normalize if checkbox is checked
        if (chkNormalizeAudio.Checked)
        {
            try
            {
                sourceFile = NormalizeWavFile(ofd.FileName, trkCompression.Value);
            }
            catch
            {
                // Use original file on failure
                sourceFile = ofd.FileName;
            }
        }

        a.AssignedSounds[tag] = sourceFile;
        lblSoundAssigned.Text = Path.GetFileName(sourceFile);
        btnRemoveSound.Visible = true;
    }

    void RemoveImageFromSelectedTag()
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) return;
        var tag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(tag)) return;

        SaveStateForUndo();
        var a = announcers[aidx];
        a.AssignedImages.Remove(tag);
        lblImageAssigned.Text = Localization.Get("NoImageAssigned");
        pbTagImage.Image = null;
        btnRemoveImage.Visible = false;
    }

    void RemoveSoundFromSelectedTag()
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) return;
        var tag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(tag)) return;

        SaveStateForUndo();
        var a = announcers[aidx];
        a.AssignedSounds.Remove(tag);
        lblSoundAssigned.Text = Localization.Get("NoSoundAssigned");
        btnRemoveSound.Visible = false;
    }


    string NormalizeWavFile(string inputPath, int compressionLevel)
    {
        // Create a temporary file for the normalized audio
        var outputPath = Path.Combine(Path.GetTempPath(), $"normalized_{Guid.NewGuid()}.wav");

        // compressionLevel: 0 = no compression (just normalize), 100 = maximum compression
        float compressionAmount = compressionLevel / 100f;

        const float targetRms = 0.22f;
        const float maxOutput = 0.95f;

        using (var reader = new AudioFileReader(inputPath))
        {
            int sampleRate = reader.WaveFormat.SampleRate;
            int channels = reader.WaveFormat.Channels;

            // Read all samples into memory for two-pass processing
            var allSamples = new List<float>();
            float[] buffer = new float[sampleRate * channels];
            int samplesRead;

            while ((samplesRead = reader.Read(buffer, 0, buffer.Length)) > 0)
            {
                for (int i = 0; i < samplesRead; i++)
                {
                    allSamples.Add(buffer[i]);
                }
            }

            if (allSamples.Count == 0)
            {
                return inputPath;
            }

            // Calculate window-based RMS for dynamic compression
            // Use ~50ms windows for smooth envelope
            int windowSize = Math.Max(1, sampleRate / 20);
            int totalWindows = (allSamples.Count + windowSize - 1) / windowSize;
            float[] windowRms = new float[totalWindows];

            for (int w = 0; w < totalWindows; w++)
            {
                int start = w * windowSize;
                int end = Math.Min(start + windowSize, allSamples.Count);
                double sum = 0;
                for (int i = start; i < end; i++)
                {
                    sum += allSamples[i] * allSamples[i];
                }
                windowRms[w] = (float)Math.Sqrt(sum / (end - start));
            }

            // Smooth the RMS envelope (moving average) - more smoothing for less artifacts
            int smoothingWindow = 7;
            float[] smoothedRms = new float[totalWindows];
            for (int w = 0; w < totalWindows; w++)
            {
                float sum = 0;
                int count = 0;
                for (int j = Math.Max(0, w - smoothingWindow); j <= Math.Min(totalWindows - 1, w + smoothingWindow); j++)
                {
                    sum += windowRms[j];
                    count++;
                }
                smoothedRms[w] = sum / count;
            }

            // Calculate overall RMS (excluding very quiet windows)
            float overallRms = 0;
            int rmsCount = 0;
            for (int w = 0; w < totalWindows; w++)
            {
                if (smoothedRms[w] > 0.01f)
                {
                    overallRms += smoothedRms[w];
                    rmsCount++;
                }
            }
            if (rmsCount > 0) overallRms /= rmsCount;

            if (overallRms < 0.005f)
            {
                return inputPath;
            }

            // Base gain to reach target RMS
            float baseGain = targetRms / overallRms;
            if (baseGain > 4f) baseGain = 4f;
            if (baseGain < 0.25f) baseGain = 0.25f;

            // Compression parameters scaled by compressionLevel
            // At 0%: no compression (ratio = 1:1)
            // At 100%: heavy compression (ratio = 6:1)
            float compRatio = 1f + (5f * compressionAmount);  // 1:1 to 6:1
            float compThreshold = targetRms * (1f - 0.5f * compressionAmount);  // Higher threshold at low compression

            // Apply compression and limiting per-sample
            float[] outputSamples = new float[allSamples.Count];

            for (int i = 0; i < allSamples.Count; i++)
            {
                int windowIdx = Math.Min(i / windowSize, totalWindows - 1);
                float localRms = smoothedRms[windowIdx];

                // Calculate dynamic gain based on local loudness
                float dynamicGain = baseGain;

                if (compressionAmount > 0.01f && localRms > 0.01f)
                {
                    float localLevel = localRms * baseGain;

                    // Apply compression if above threshold
                    if (localLevel > compThreshold)
                    {
                        float excess = localLevel - compThreshold;
                        float compressedExcess = excess / compRatio;
                        float targetLevel = compThreshold + compressedExcess;
                        dynamicGain = baseGain * (targetLevel / localLevel);
                    }
                }

                float sample = allSamples[i] * dynamicGain;

                // Soft limiting using smooth curve
                float absSample = Math.Abs(sample);
                if (absSample > 0.8f)
                {
                    // Gentler soft knee limiting
                    float x = (absSample - 0.8f) / 0.2f;
                    float limited = 0.8f + 0.15f * x / (1f + x);
                    if (limited > maxOutput) limited = maxOutput;
                    sample = sample > 0 ? limited : -limited;
                }

                outputSamples[i] = sample;
            }

            // Write output
            using (var writer = new WaveFileWriter(outputPath, reader.WaveFormat))
            {
                // Write in chunks
                int chunkSize = sampleRate * channels;
                for (int i = 0; i < outputSamples.Length; i += chunkSize)
                {
                    int count = Math.Min(chunkSize, outputSamples.Length - i);
                    float[] chunk = new float[count];
                    Array.Copy(outputSamples, i, chunk, 0, count);
                    writer.WriteSamples(chunk, 0, count);
                }
            }
        }

        return outputPath;
    }


    void PlayAssignedSound()
    {
        var aidx = lstAnnouncers.SelectedIndex;
        if (aidx < 0) return;
        var tag = lstTags.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(tag)) return;
        
        var a = announcers[aidx];
        if (!a.AssignedSounds.TryGetValue(tag, out var soundPath) || string.IsNullOrEmpty(soundPath) || !File.Exists(soundPath))
        {
            MessageBox.Show(Localization.Get("NoSoundAssigned"), Localization.Get("PlaySound"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        
        try
        {
            // Stop any currently playing audio
            StopSound();
            
            // Play the sound
            audioReader = new AudioFileReader(soundPath);
            waveOut = new WaveOutEvent();
            waveOut.Init(audioReader);
            waveOut.PlaybackStopped += (s, e) => StopSound();
            waveOut.Play();
        }
        catch (Exception ex)
        {
            MessageBox.Show(Localization.Get("PlaySoundFailed") + ex.Message, Localization.Get("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    void StopSound()
    {
        if (waveOut != null)
        {
            waveOut.Stop();
            waveOut.Dispose();
            waveOut = null;
        }
        if (audioReader != null)
        {
            audioReader.Dispose();
            audioReader = null;
        }
    }

    void SelectSaveFolder()
    {
        using var fbd = new FolderBrowserDialog();
        fbd.Description = "Select folder where the announcer mod will be saved (the folder itself will contain AnnouncersImage_LEB etc.)";
        if (fbd.ShowDialog() != DialogResult.OK) return;
        saveFolder = fbd.SelectedPath;
        lblFolder.Text = saveFolder;
    }

    void SaveMod()
    {
        if (string.IsNullOrEmpty(saveFolder)) { MessageBox.Show(Localization.Get("SelectFolderFirst")); return; }
        try
        {
            // create root subfolders
            var imgRoot = Path.Combine(saveFolder, "AnnouncersImage_LEB");
            var txtRoot = Path.Combine(saveFolder, "AnnouncersTEXT_LEB");
            var sndRoot = Path.Combine(saveFolder, "AnnouncersSounds_LEB");
            Directory.CreateDirectory(imgRoot);
            Directory.CreateDirectory(txtRoot);
            Directory.CreateDirectory(sndRoot);

            // build AnnouncersXML_LEB.xml
            var xmlRoot = new XElement("AnnouncersXML");
            foreach(var a in announcers)
            {
                var imgDest = Path.Combine(imgRoot, a.Name);
                Directory.CreateDirectory(imgDest);

                // copy Announcer.png
                if (!string.IsNullOrEmpty(a.AnnouncerImage) && File.Exists(a.AnnouncerImage))
                {
                    File.Copy(a.AnnouncerImage, Path.Combine(imgDest, "Announcer.png"), true);
                }

                // copy UI.png
                if (!string.IsNullOrEmpty(a.UIImage) && File.Exists(a.UIImage))
                {
                    File.Copy(a.UIImage, Path.Combine(imgDest, "UI.png"), true);
                }

                var ann = new XElement("Announcers",
                    new XElement("Name", a.Name),
                    new XElement("Normal_A_Value", "true"),
                    new XElement("R", a.BorderColor.R),
                    new XElement("G", a.BorderColor.G),
                    new XElement("B", a.BorderColor.B),
                    new XElement("A", a.BorderColor.A)
                );
                if (a.Random) ann.Add(new XElement("Random", a.Random.ToString().ToLower()));
                if (a.Quantity != 1) ann.Add(new XElement("Quantity", a.Quantity));
                // Always output Expression as true
                ann.Add(new XElement("Expression", "true"));
                xmlRoot.Add(ann);

                // copy assigned images
                foreach(var kv in a.AssignedImages)
                {
                    var tag = kv.Key;
                    var src = kv.Value;
                    if (!File.Exists(src)) continue;
                    var baseName = tag.Replace("_TEXT", "").Trim('_');
                    var destName = $"Announcer{baseName}.png";
                    var dest = Path.Combine(imgDest, destName);
                    File.Copy(src, dest, true);
                }

                // copy assigned sounds to AnnouncersSounds_LEB folder
                var sndDest = Path.Combine(sndRoot, a.Name);
                Directory.CreateDirectory(sndDest);
                foreach(var kv in a.AssignedSounds)
                {
                    var tag = kv.Key;
                    var src = kv.Value;
                    if (!File.Exists(src)) continue;
                    // Format: {TagName}.wav (e.g., AgentDie.wav, AgentDie_1.wav)
                    var destName = $"{tag}.wav";
                    var dest = Path.Combine(sndDest, destName);
                    File.Copy(src, dest, true);
                }

                // texts per language
                var announcerTextDir = Path.Combine(txtRoot, a.Name);
                Directory.CreateDirectory(announcerTextDir);
                foreach(var lang in a.Texts.Keys)
                {
                    var langDir = Path.Combine(announcerTextDir, lang);
                    Directory.CreateDirectory(langDir);
                    var xmlPath = Path.Combine(langDir, lang + ".xml");
                    var langRoot = new XElement("Language");
                    if (a.Texts.TryGetValue(lang, out var texts))
                    {
                        // Sort tags: first by base tag order (Tags array), then by suffix number
                        var sortedTexts = texts
                            .Where(kv => !string.IsNullOrEmpty(kv.Value))
                            .OrderBy(kv => {
                                string baseTag = GetBaseTag(kv.Key);
                                int index = Array.IndexOf(Tags, baseTag);
                                return index >= 0 ? index : int.MaxValue;
                            })
                            .ThenBy(kv => {
                                if (kv.Key.Contains("_") && int.TryParse(kv.Key.Split('_').Last(), out var n))
                                    return n;
                                return 0;
                            });

                        foreach(var kv in sortedTexts)
                        {
                            // Add _TEXT suffix for XML element names
                            // Format: {BaseTag}_TEXT or {BaseTag}_TEXT_{Number}
                            // Internal key is like "AgentDie" or "AgentDie_1"
                            // Output should be "AgentDie_TEXT" or "AgentDie_TEXT_1"
                            var xmlTagName = ConvertToXmlTagName(kv.Key);
                            langRoot.Add(new XElement(xmlTagName, kv.Value));
                        }
                    }
                    // Write XML with blank lines between different base tags
                    using (var sw = new StreamWriter(xmlPath, false, System.Text.Encoding.UTF8))
                    {
                        sw.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>");
                        sw.WriteLine("<Language>");
                        string? lastBaseTag = null;
                        foreach (var el in langRoot.Elements())
                        {
                            // Extract base tag from XML element name (e.g., "AgentDie_TEXT_1" -> "AgentDie")
                            var elName = el.Name.LocalName;
                            var currentBaseTag = elName.Replace("_TEXT", "");
                            if (currentBaseTag.Contains("_") && int.TryParse(currentBaseTag.Split('_').Last(), out _))
                            {
                                var parts = currentBaseTag.Split('_');
                                currentBaseTag = string.Join("_", parts.Take(parts.Length - 1));
                            }

                            // Add blank line when base tag changes (except for first element)
                            if (lastBaseTag != null && lastBaseTag != currentBaseTag)
                            {
                                sw.WriteLine();
                            }
                            lastBaseTag = currentBaseTag;
                            sw.WriteLine($"  <{el.Name}>{el.Value}</{el.Name}>");
                        }
                        sw.WriteLine("</Language>");
                    }
                }
            }

            var xmlDoc = new XDocument(new XDeclaration("1.0","utf-8","yes"), xmlRoot);
            xmlDoc.Save(Path.Combine(saveFolder, "AnnouncersXML_LEB.xml"));

            MessageBox.Show(Localization.Get("ModSavedSuccess"));
        }
        catch(Exception ex)
        {
            MessageBox.Show(Localization.Get("ErrorWhileSaving") + ex.Message);
        }
    }

    void LoadExistingMod()
    {
        using var ofd = new OpenFileDialog();
        ofd.Filter = "AnnouncersXML_LEB.xml|AnnouncersXML_LEB.xml|All files|*.*";
        ofd.Title = Localization.Get("SelectXmlToLoad");
        if (ofd.ShowDialog() != DialogResult.OK) return;
        var xmlPath = ofd.FileName;
        var baseDir = Path.GetDirectoryName(xmlPath) ?? "";
        var imageRoot = Path.Combine(baseDir, "AnnouncersImage_LEB");
        var textRoot = Path.Combine(baseDir, "AnnouncersTEXT_LEB");
        try
        {
            var doc = XDocument.Load(xmlPath);
            var root = doc.Root;
            if (root == null) throw new Exception("Invalid XML");
            announcers.Clear(); lstAnnouncers.Items.Clear();
            foreach(var annElem in root.Elements("Announcers"))
            {
                var name = annElem.Element("Name")?.Value ?? "NewAnnouncer";
                var a = new Announcer() { Name = name };
                var r = int.TryParse(annElem.Element("R")?.Value, out var rv) ? rv : 172;
                var g = int.TryParse(annElem.Element("G")?.Value, out var gv) ? gv : 219;
                var b = int.TryParse(annElem.Element("B")?.Value, out var bv) ? bv : 242;
                // Check if Normal_A_Value is true (0-255 range) or false/missing (0-1 range)
                var normalAValue = annElem.Element("Normal_A_Value")?.Value == "true";
                int alpha;
                if (normalAValue)
                {
                    alpha = int.TryParse(annElem.Element("A")?.Value, out var av) ? av : 255;
                }
                else
                {
                    alpha = double.TryParse(annElem.Element("A")?.Value, out var av) ? (int)(av * 255) : 255;
                }
                a.BorderColor = Color.FromArgb(alpha, r, g, b);
                a.Random = (annElem.Element("Random")?.Value == "true");
                a.Expression = true; // Always true
                a.Quantity = int.TryParse(annElem.Element("Quantity")?.Value, out var qv) ? qv : 1;

                // load texts if available
                var announcerTextDir = Path.Combine(textRoot, a.Name);
                if (Directory.Exists(announcerTextDir))
                {
                    foreach(var langDir in Directory.GetDirectories(announcerTextDir))
                    {
                        var lang = Path.GetFileName(langDir);
                        var xmlFile = Path.Combine(langDir, lang + ".xml");
                        if (!File.Exists(xmlFile)) continue;
                        var docLang = XDocument.Load(xmlFile);
                        var langRoot = docLang.Root;
                        if (langRoot == null) continue;
                        var map = new Dictionary<string, string>();
                        foreach(var el in langRoot.Elements())
                        {
                            // Remove _TEXT suffix from XML element names for internal storage
                            var tagName = el.Name.LocalName.Replace("_TEXT", "");
                            map[tagName] = el.Value;
                        }
                        a.Texts[lang] = map;
                    }
                }

                // find assigned images where present
                var imgDir = Path.Combine(imageRoot, a.Name);
                if (Directory.Exists(imgDir))
                {
                    var pngFiles = Directory.GetFiles(imgDir, "Announcer*.png");
                    foreach(var fname in pngFiles)
                    {
                        var baseName = Path.GetFileNameWithoutExtension(fname).Replace("Announcer", "");
                        // Internal tags no longer use _TEXT suffix
                        a.AssignedImages[baseName] = fname;
                    }
                    // load Announcer.png and UI.png
                    var announcerPng = Path.Combine(imgDir, "Announcer.png");
                    if (File.Exists(announcerPng)) a.AnnouncerImage = announcerPng;
                    var uiPng = Path.Combine(imgDir, "UI.png");
                    if (File.Exists(uiPng)) a.UIImage = uiPng;

                }

                // find assigned sounds in AnnouncersSounds_LEB folder
                var sndRoot = Path.Combine(baseDir, "AnnouncersSounds_LEB");
                var sndDir = Path.Combine(sndRoot, a.Name);
                if (Directory.Exists(sndDir))
                {
                    var wavFiles = Directory.GetFiles(sndDir, "*.wav");
                    foreach(var fname in wavFiles)
                    {
                        // Format: {TagName}.wav (e.g., AgentDie.wav, AgentDie_1.wav)
                        var tagName = Path.GetFileNameWithoutExtension(fname);
                        a.AssignedSounds[tagName] = fname;
                    }
                }

                // If texts exist, set AutoGenerateTags to false (manual mode)
                if (a.Texts.Count > 0 && a.Texts.Values.Any(t => t.Count > 0))
                {
                    a.AutoGenerateTags = false;
                }

                announcers.Add(a);
                lstAnnouncers.Items.Add(a.Name);
            }
            if (lstAnnouncers.Items.Count>0) lstAnnouncers.SelectedIndex = 0;
            // update language combo with languages from loaded announcers only
            cmbLanguage.Items.Clear();
            foreach(var a in announcers)
            {
                foreach(var lang in a.Texts.Keys)
                {
                    if (!cmbLanguage.Items.Contains(lang)) cmbLanguage.Items.Add(lang);
                }
            }
            // Select first available language if any
            if (cmbLanguage.Items.Count > 0)
                cmbLanguage.SelectedIndex = 0;
            MessageBox.Show(Localization.Get("ModLoadedSuccess"));
        }
        catch(Exception ex)
        {
            MessageBox.Show(Localization.Get("FailedToLoadMod") + ex.Message);
        }
    }

    void AssignAnnouncerImage()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) return;
        var a = announcers[idx];
        using (var ofd = new OpenFileDialog() { Filter = "PNG files|*.png" })
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                // Validate and optionally resize image (required: 512x512)
                var imagePath = ShowResizeDialog(ofd.FileName, 512, 512, "Announcer.png");
                if (imagePath == null) return;

                SaveStateForUndo();
                a.AnnouncerImage = imagePath;
                lblAnnouncerImage.Text = Path.GetFileName(imagePath);
                pbAnnouncerImage.Image = Image.FromFile(imagePath);
            }
        }
    }

    void AssignUIImage()
    {
        var idx = lstAnnouncers.SelectedIndex;
        if (idx < 0) return;
        var a = announcers[idx];
        using (var ofd = new OpenFileDialog() { Filter = "PNG files|*.png" })
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                // Validate and optionally resize image (recommended: 345x213)
                var imagePath = ShowResizeDialog(ofd.FileName, 345, 213, "UI.png");
                if (imagePath == null) return;

                SaveStateForUndo();
                a.UIImage = imagePath;
                lblUIImage.Text = Path.GetFileName(imagePath);
                pbUIImage.Image = Image.FromFile(imagePath);
            }
        }
    }

    /// <summary>
    /// Shows a dialog to choose resize method and returns the resized image path, or null if cancelled.
    /// </summary>
    string? ShowResizeDialog(string originalPath, int targetWidth, int targetHeight, string imageType)
    {
        using var img = Image.FromFile(originalPath);

        if (img.Width == targetWidth && img.Height == targetHeight)
        {
            return originalPath; // No resize needed
        }

        var result = MessageBox.Show(
            string.Format(Localization.Get("ImageSizeWarning"), img.Width, img.Height, imageType, targetWidth, targetHeight),
            Localization.Get("ImageSizeMismatch"),
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (result == DialogResult.Cancel)
        {
            return originalPath; // Use original
        }

        // Create resized image
        var resizedPath = Path.Combine(Path.GetTempPath(), $"resized_{Guid.NewGuid()}.png");

        if (result == DialogResult.Yes)
        {
            // Scale method
            ResizeImageScale(originalPath, resizedPath, targetWidth, targetHeight);
        }
        else
        {
            // Crop method
            ResizeImageCrop(originalPath, resizedPath, targetWidth, targetHeight);
        }

        return resizedPath;
    }

    /// <summary>
    /// Resizes image by scaling (stretching/shrinking to fit target size)
    /// </summary>
    void ResizeImageScale(string sourcePath, string destPath, int targetWidth, int targetHeight)
    {
        using var original = Image.FromFile(sourcePath);
        using var resized = new Bitmap(targetWidth, targetHeight);
        using var g = Graphics.FromImage(resized);

        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

        g.DrawImage(original, 0, 0, targetWidth, targetHeight);
        resized.Save(destPath, System.Drawing.Imaging.ImageFormat.Png);
    }

    /// <summary>
    /// Resizes image by cropping from center
    /// </summary>
    void ResizeImageCrop(string sourcePath, string destPath, int targetWidth, int targetHeight)
    {
        using var original = Image.FromFile(sourcePath);
        using var resized = new Bitmap(targetWidth, targetHeight);
        using var g = Graphics.FromImage(resized);

        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

        // Calculate scale to cover target area
        float scaleX = (float)targetWidth / original.Width;
        float scaleY = (float)targetHeight / original.Height;
        float scale = Math.Max(scaleX, scaleY);

        int scaledWidth = (int)(original.Width * scale);
        int scaledHeight = (int)(original.Height * scale);

        // Center the image
        int offsetX = (targetWidth - scaledWidth) / 2;
        int offsetY = (targetHeight - scaledHeight) / 2;

        g.DrawImage(original, offsetX, offsetY, scaledWidth, scaledHeight);
        resized.Save(destPath, System.Drawing.Imaging.ImageFormat.Png);
    }

    static void CreatePlaceholderPng(string path, string label)
    {
        using var bmp = new Bitmap(512, 512);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.FromArgb(255, 50, 50, 50));
        var pen = Pens.Gray;
        g.DrawRectangle(pen, 10, 10, 492, 492);
        var font = new Font("Segoe UI", 20);
        var brush = Brushes.White;
        var text = Path.GetFileNameWithoutExtension(label);
        var sf = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(text, font, brush, new RectangleF(0, 0, 512, 512), sf);
        bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    }
}
