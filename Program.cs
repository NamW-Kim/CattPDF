using Spire.Pdf;
using Spire.Pdf.Fields;
using Spire.Pdf.Widget;
using System.Data;
using System.Diagnostics;

// This loads data portion of the file
DataTable datatable = new DataTable();
StreamReader streamreader = new StreamReader("PDF/data.txt");
char[] delimiter = new char[] { '\t' };
string[] columnheaders = streamreader.ReadLine().Split(delimiter);
foreach (string columnheader in columnheaders)
{
    datatable.Columns.Add(columnheader);
}

streamreader = new StreamReader("PDF/data.txt"); // RESET COUNT

while (streamreader.Peek() > 0)
{
    DataRow datarow = datatable.NewRow();
    datarow.ItemArray = streamreader.ReadLine().Split(delimiter);
    datatable.Rows.Add(datarow);
}

// TEST PRINT OF 2D ARRAY DATATABLE
foreach (DataRow row in datatable.Rows)
{
    Console.WriteLine("----Row No: " + datatable.Rows.IndexOf(row) + "----");
    foreach (DataColumn column in datatable.Columns)
    {
        Console.WriteLine("Index: "+datatable.Columns.IndexOf(column));
        Console.WriteLine(row[column]);
    }
}

// What does this do? It makes CMD line halt
// Console.ReadLine();

// -------------------------------------------------------------------------------

// This loads PDF reader
PdfDocument doc = new PdfDocument();
doc.LoadFromFile("PDF/BOL.pdf");
PdfFormWidget formWidget = doc.Form as PdfFormWidget;


foreach (DataRow row in datatable.Rows)
{
    foreach (DataColumn column in datatable.Columns)
    {
        
        // Write to PDF
        for (int i = 0; i < formWidget.FieldsWidget.List.Count; i++)
        {
            PdfField field = formWidget.FieldsWidget.List[i] as PdfField;

            string fieldName = field.Name;
            //Console.WriteLine(fieldName);
            //Console.WriteLine(field.Name.GetType());
            if (field is PdfTextBoxFieldWidget)
            {
                PdfTextBoxFieldWidget textBoxField = field as PdfTextBoxFieldWidget;
                switch (textBoxField.Name)
                {
                    case "date":
                        textBoxField.Text = DateTime.Today.ToString("d");
                        break;
                    case "stname": // SHIP TO NAME
                        textBoxField.Text = "Target DC " + row[datatable.Columns[1]];
                        break;
                    case "stadd": // SHIP TO ADDRESS 1ST LINE
                        textBoxField.Text = row[datatable.Columns[3]].ToString();
                        break;
                    case "stcsz": // SHIP TO CITY/STATE/ZIP FIELD
                        textBoxField.Text = row[datatable.Columns[4]].ToString()+", "+ row[datatable.Columns[5]].ToString()+" "+ Convert.ToInt32(row[datatable.Columns[6]]).ToString("00000");
                        break;
                    case "con1": // CUSTOMER ORDER INFO: CUSTOMER ORDER NUMBER
                        textBoxField.Text = row[datatable.Columns[0]].ToString();
                        break;
                    case "pkgs1": // CUSTOMER ORDER INFO: NUMBER OF PKGS
                        textBoxField.Text = row[datatable.Columns[17]].ToString();
                        break;
                    case "pkgstot": // CUSTOMER ORDER INFO: NUMBER OF PKGS TOTAL
                        textBoxField.Text = row[datatable.Columns[17]].ToString();
                        break;
                    case "wgt1": //CUSTOMER ORDER INFO: WEIGHT
                        textBoxField.Text = row[datatable.Columns[18]].ToString();
                        break;
                    case "wgttot": //CUSTOMER ORDER INFO: WEIGHT
                        textBoxField.Text = row[datatable.Columns[18]].ToString();
                        break;
                    case "psy1": //CUSTOMER ORDER INFO: PALLET SLIP (CHECK IF THIS IS FEDEX OR LTL)
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("ltl"))
                        {
                            textBoxField.Text = "Y";
                        }
                        else
                        {
                            textBoxField.Text = "";
                        }
                        break;
                    case "psn1": //CUSTOMER ORDER INFO: PALLET SLIP (CHECK IF THIS IS FEDEX OR LTL)
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("fedex"))
                        {
                            textBoxField.Text = "N";
                        }
                        else
                        {
                            textBoxField.Text = "";
                        }
                        break;

                    // -------------------------------------------------------------------

                    case "huq1": // CARRIER INFO: HANDLING UNIT: QTY
                        textBoxField.Text = row[datatable.Columns[17]].ToString();
                        break;
                    case "huqtot": // CARRIER INFO: HANDLING UNIT: QTY TOTAL
                        textBoxField.Text = row[datatable.Columns[17]].ToString();
                        break;
                    case "hut1": // CARRIER INFO: HANDLING UNIT: TYPE
                        textBoxField.Text = "Box";
                        break;
                    case "pq1": // CARRIER INFO: PACKAGE: QTY
                        textBoxField.Text = row[datatable.Columns[16]].ToString();
                        break;
                    case "pqtot": // CARRIER INFO: PACKAGE: QTY TOTAL
                        textBoxField.Text = row[datatable.Columns[16]].ToString();
                        break;
                    case "pt1": // CARRIER INFO: PACKAGE: TYPE
                        textBoxField.Text = "Units";
                        break;
                    case "ciw1": //CARRIER INFO:: WEIGHT
                        textBoxField.Text = row[datatable.Columns[18]].ToString();
                        break;
                    case "ciwtot": //CARRIER INFO:: WEIGHT
                        textBoxField.Text = row[datatable.Columns[18]].ToString();
                        break;
                    case "cd1": //CARRIER INFO: COMMODITY DESCRIPTION
                        textBoxField.Text = "Laundry Detergent Strips";
                        break;

                    // -------------------------------------------------------------------

                    case "boln": //BOL NUMBER
                        textBoxField.Text = row[datatable.Columns[2]].ToString();
                        break;
                    case "cn": //CARRIER NAME: FEDEX OR RDWY
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("fedex"))
                        {
                            textBoxField.Text = "FEDEX";
                        }
                        else
                        {
                            textBoxField.Text = "RDWY";
                        }
                        break;
                    case "scac": //SCAC NAME: FEDEX (FDEG OR RDWY)
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("fedex"))
                        {
                            textBoxField.Text = "FDEG";
                        }
                        else
                        {
                            textBoxField.Text = "RDWY";
                        }
                        break;
                }
            }

            /*
            if (field is PdfListBoxWidgetFieldWidget)
            {
                PdfListBoxWidgetFieldWidget listBoxField = field as PdfListBoxWidgetFieldWidget;
                switch (listBoxField.Name)
                {
                    case "email_format":
                        int[] index = { 1 };
                        listBoxField.SelectedIndex = index;
                        break;
                }
            }

            if (field is PdfComboBoxWidgetFieldWidget)
            {
                PdfComboBoxWidgetFieldWidget comBoxField = field as PdfComboBoxWidgetFieldWidget;
                switch (comBoxField.Name)
                {
                    case "title":
                        int[] items = { 0 };
                        comBoxField.SelectedIndex = items;
                        break;
                }
            }

            if (field is PdfRadioButtonListFieldWidget)
            {
                PdfRadioButtonListFieldWidget radioBtnField = field as PdfRadioButtonListFieldWidget;
                switch (radioBtnField.Name)
                {
                    case "country":
                        radioBtnField.SelectedIndex = 1;
                        break;
                }
            }

            if (field is PdfCheckBoxWidgetFieldWidget)
            {
                PdfCheckBoxWidgetFieldWidget checkBoxField = field as PdfCheckBoxWidgetFieldWidget;
                switch (checkBoxField.Name)
                {
                    case "agreement_of_terms":
                        checkBoxField.Checked = true;
                        break;
                }
            }
            if (field is PdfButtonWidgetFieldWidget)
            {
                PdfButtonWidgetFieldWidget btnField = field as PdfButtonWidgetFieldWidget;
                switch (btnField.Name)
                {
                    case "submit":
                        btnField.Text = "Submit";
                        break;
                }
            }

            */
        }
    }
    // Export to PDF
    doc.SaveToFile("Export/BOL Target "+ row[datatable.Columns[2]].ToString() +" "+ DateTime.Today.ToString("d") + ".pdf");

}
// Process.Start("explorer.exe", "/select, c:\\teste");
