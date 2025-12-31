using QuestPDF.Infrastructure;
using SnackOfficeApp;

ApplicationConfiguration.Initialize();

QuestPDF.Settings.License = LicenseType.Community;

try
{
    AppDb.Initialize();
}
catch (Exception ex)
{
    MessageBox.Show(
        $"Database init failed:\n{ex.Message}",
        "SnackOfficeApp",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
    );
    return;
}

Application.Run(new MainForm());
