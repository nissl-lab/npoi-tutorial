using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using (var wb = new XSSFWorkbook())
{
    var sheet = wb.CreateSheet("Pictures");
    var drawing = sheet.CreateDrawingPatriarch();
    var anchor = new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 5);
    anchor.AnchorType = AnchorType.MoveDontResize;
    int imageId = LoadPNGImage("dotnet.png", wb);
    var pic= drawing.CreatePicture(anchor, imageId);

    using (var file = File.Create("pictures.xlsx"))
    {
        wb.Write(file);
    }
}

static int LoadPNGImage(string path, IWorkbook wb)
{
    using (FileStream file = File.OpenRead(path))
    {
        byte[] buffer =new byte[file.Length];
        file.Read(buffer, 0, (int)file.Length);
        return wb.AddPicture(buffer, PictureType.PNG);
    }
}