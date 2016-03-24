using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    private DataSet ds1, ds2;
    private DataTable dt1, dt2, dtResult;
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    private DataSet GetDSFromExcel(string path)
    {
        DataSet ds = new DataSet();

        var file = new FileInfo(path);
        using (var stream = new FileStream(path, FileMode.Open))
        {
            IExcelDataReader reader = null;
            if (file.Extension == ".xls")
            {
                reader = ExcelReaderFactory.CreateBinaryReader(stream);

            }
            else if (file.Extension == ".xlsx")
            {
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            if (reader == null) { return null; }

            reader.IsFirstRowAsColumnNames = true;
            ds = reader.AsDataSet();
        }
        return ds;
    }
    private DataTable CompareDT(DataTable TableA, DataTable TableB)
    {
        DataTable dt = new DataTable();

        var results = from table1 in TableA.AsEnumerable()
                      join table2 in TableB.AsEnumerable() on table1.Field<string>("RowID") equals table2.Field<string>("RowID")
                      where table1.Field<String>("Pharmacy Item Description") != table2.Field<String>("Pharmacy Item Description")
                      select table1;
        //This will give you the rows in dt1 which do not match the rows in dt2.  You will need to expand the where clause to include all your columns.


        //If you do not have primarry keys then you will need to match up each column and then find the missing.
        var matched = from table1 in TableA.AsEnumerable()
                      join table2 in TableB.AsEnumerable() on table1.Field<string>("RowID") equals table2.Field<string>("RowID")
                      where table1.Field<string>("Pharmacy Item Description") == table2.Field<string>("Pharmacy Item Description") 
                      select table1;
        var missing = from table1 in TableA.AsEnumerable()
                      where !matched.Contains(table1)
                      select table1;

        return dt;
    }
    protected void btnCompare_Click(object sender, EventArgs e)
    {
        SaveFile(File1.PostedFile);
        SaveFile(File2.PostedFile);
        //Label1.Text = File1.PostedFile.FileName;
        var path1 = Server.MapPath("file\\" + File1.FileName);
        var path2 = Server.MapPath("file\\" + File2.FileName);

        ds1 = GetDSFromExcel(path1);
        //ds2 = GetDSFromExcel(path2);

        var file = new FileInfo(path2);
        using (var stream = new FileStream(path2, FileMode.Open))
        {
            IExcelDataReader reader = null;
            if (file.Extension == ".xls")
            {
                reader = ExcelReaderFactory.CreateBinaryReader(stream);

            }
            else if (file.Extension == ".xlsx")
            {
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            if (reader == null) { return; }

            reader.IsFirstRowAsColumnNames = true;
            ds2 = reader.AsDataSet();
        }

        dtResult = CompareDT(ds1.Tables[0], ds2.Tables[0]);
    }
    void SaveFile(HttpPostedFile file)
    {
        // Specify the path to save the uploaded file to.
        string savePath = Server.MapPath("file\\");

        // Get the name of the file to upload.
        string fileName = file.FileName;

        // Create the path and file name to check for duplicates.
        string pathToCheck = savePath + fileName;

        // Create a temporary file name to use for checking duplicates.
        string tempfileName = "";

        // Check to see if a file already exists with the
        // same name as the file to upload.        
        if (System.IO.File.Exists(pathToCheck))
        {
            int counter = 2;
            while (System.IO.File.Exists(pathToCheck))
            {
                // if a file with this name already exists,
                // prefix the filename with a number.
                tempfileName = counter.ToString() + fileName;
                pathToCheck = savePath + tempfileName;
                counter++;
            }

            fileName = tempfileName;
        }
        else
        {
            // Notify the user that the file was saved successfully.
        }

        // Append the name of the file to upload to the path.
        savePath += fileName;

        // Call the SaveAs method to save the uploaded
        // file to the specified directory.
        File1.SaveAs(savePath);

    }
}