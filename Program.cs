using Spire.Pdf;
using Spire.Pdf.Fields;
using Spire.Pdf.Widget;
using System.Data;
using System.Diagnostics;

// This loads data portion of the file
DataTable datatable = new DataTable();
StreamReader streamreader = new StreamReader(AppContext.BaseDirectory + "PDF/data.txt");
char[] delimiter = new char[] { '\t' };
string[] columnheaders = streamreader.ReadLine().Split(delimiter);
foreach (string columnheader in columnheaders)
{
    datatable.Columns.Add(columnheader);
}

streamreader = new StreamReader(AppContext.BaseDirectory + "PDF/data.txt"); // RESET COUNT

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
PdfDocument BOL_DOCUMENT = new PdfDocument();
PdfDocument CUSTOM_DOCUMENT = new PdfDocument();
BOL_DOCUMENT.LoadFromFile(AppContext.BaseDirectory+"PDF/BOL.pdf");
CUSTOM_DOCUMENT.LoadFromFile(AppContext.BaseDirectory + "PDF/CUSTOM.pdf");
PdfFormWidget? formWidgetBOL = BOL_DOCUMENT.Form as PdfFormWidget;
PdfFormWidget? formWidgetCUSTOM = CUSTOM_DOCUMENT.Form as PdfFormWidget;

// THIS IS ENTIRELY FOR BOL GENERATION
foreach (DataRow row in datatable.Rows)
{
    foreach (DataColumn column in datatable.Columns)
    {
        
        // Write to PDF
        for (int i = 0; i < formWidgetBOL.FieldsWidget.List.Count; i++)
        {
            PdfField field = formWidgetBOL.FieldsWidget.List[i] as PdfField;

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
                    case "class1": //LTL ONLY: First Line of Class
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("fedex"))
                        {
                            textBoxField.Text = "";
                        }
                        else
                        {
                            textBoxField.Text = "55";
                        }
                        break;
                    case "asi1": //CUSTOMER ORDER INFORMATION: ADDITIONAL SHIPPER INFO
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("fedex"))
                        {
                            textBoxField.Text = "Individual Carton(s)";
                        }
                        else
                        {
                            textBoxField.Text = "Pure Pallet(s)";
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
    BOL_DOCUMENT.SaveToFile(AppContext.BaseDirectory + "Export/BOL Target " + row[datatable.Columns[1]].ToString() +" "+ DateTime.Today.ToString("d") + ".pdf");

}

// THIS IS ENTIRELY FOR CUSTOM GENERATION
foreach (DataRow row in datatable.Rows)
{
    foreach (DataColumn column in datatable.Columns)
    {

        // Write to PDF
        for (int i = 0; i < formWidgetCUSTOM.FieldsWidget.List.Count; i++)
        {
            PdfField field = formWidgetCUSTOM.FieldsWidget.List[i] as PdfField;

            string fieldName = field.Name;
            //Console.WriteLine(fieldName);
            //Console.WriteLine(field.Name.GetType());
            if (field is PdfTextBoxFieldWidget)
            {
                PdfTextBoxFieldWidget textBoxField = field as PdfTextBoxFieldWidget;
                switch (textBoxField.Name)
                {
                    case "23 Date":
                        textBoxField.Text = DateTime.Today.ToString("d");
                        break;
                    case "4 Consignee Name": // SHIP TO NAME
                        textBoxField.Text = "Target DC " + row[datatable.Columns[1]];
                        break;
                    case "4 Consignee Address": // SHIP TO ADDRESS 1ST LINE + 2ND LINE
                        textBoxField.Text = row[datatable.Columns[3]].ToString() + "\n" + row[datatable.Columns[4]].ToString() + ", " + row[datatable.Columns[5]].ToString() + " " + Convert.ToInt32(row[datatable.Columns[6]]).ToString("00000");
                        break;
                    case "5 Buyer Name": // Hard Code BUYER INFO
                        textBoxField.Text = "Target Corp C/O CHRLTL";
                        break;
                    case "3 Other Ref Nos": // CUSTOMER ORDER INFO: CUSTOMER ORDER NUMBER
                        textBoxField.Text = "BOL:\n "+row[datatable.Columns[20]].ToString() + "\n \n PO:\n " + row[datatable.Columns[0]].ToString();
                        break;
                    case "18_1": // CUSTOMER ORDER INFO: NUMBER OF PKGS
                        textBoxField.Text = row[datatable.Columns[17]].ToString();
                        break;
                    case "24 Total Packages": // CUSTOMER ORDER INFO: NUMBER OF PKGS TOTAL
                        textBoxField.Text = row[datatable.Columns[17]].ToString();
                        break;
                    case "20Grs_1": //CUSTOMER ORDER INFO: WEIGHT GROSS
                        textBoxField.Text = row[datatable.Columns[18]].ToString();
                        break;
                    case "Gross Shipping Weight": //CUSTOMER ORDER INFO: WEIGHT GROSS
                        textBoxField.Text = row[datatable.Columns[18]].ToString();
                        break;
                    case "20Net_1": //CUSTOMER ORDER INFO: WEIGHT NET
                        textBoxField.Text = "10.7";
                        break;
                    case "22_1": // UNIT PRICE
                        textBoxField.Text = "7.38";
                        break;
                    case "23_1": // CARRIER INFO: HANDLING UNIT: QTY TOTAL
                        textBoxField.Text = "$"+(Convert.ToInt32(row[datatable.Columns[16]])*7.38).ToString("#,##0.00");
                        break;
                    case "26total": // CARRIER INFO: HANDLING UNIT: QTY TOTAL
                        textBoxField.Text = "$" + (Convert.ToInt32(row[datatable.Columns[16]]) * 7.38).ToString("#,##0.00");
                        break;
                    case "21UOM_1": // CARRIER INFO: HANDLING UNIT: EOM
                        textBoxField.Text = "EACH";
                        break;
                    case "21Units_1": // CARRIER INFO: PACKAGE: QTY
                        textBoxField.Text = row[datatable.Columns[16]].ToString();
                        break;
                    case "2 Exporter Name": // EXPORTER NAME
                        textBoxField.Text = "Tru Earth Environmental Products Inc.";
                        break;
                    case "2 Exporter Contact": // EXPORTER CONTACT
                        textBoxField.Text = "Jeerus Singla // Catt Kim";
                        break;
                    case "2 Exporter Phone": // EXPORTER CONTACT
                        textBoxField.Text = "(Jeerus's Phone #) // 7788892043";
                        break;
                    case "2 Exporter Address": // EXPORTER CONTACT
                        textBoxField.Text = "7500 Winston Street, Unit 108\nBurnaby, BC V5A 4J8";
                        break;
                    case "4 Consignee IRS EIN SSN": // CONSIGNEE IRS EIN SSN
                        textBoxField.Text = "41-0215170";
                        break;
                    case "8 ORIGIN COUNTRYPROVINCE": // EXPORTER NAME
                        textBoxField.Text = "CA/NB";
                        break;
                    case "9 DESTINATION COUNTRYSTATE": // EXPORTER CONTACT
                        textBoxField.Text = "US/CA";
                        break;
                    case "13 Terms of Sale Payment and Discount": // EXPORTER CONTACT
                        textBoxField.Text = "Collect";
                        break;
                    case "14 Currency Used": // CURRENCY USED
                        textBoxField.Text = "USD";
                        break;
                    case "19_1": // DESCRIPTION OF GOODS
                        textBoxField.Text = "Laundry Detergent Strips";
                        break;
                    case "17_1": // HS CODE
                        textBoxField.Text = "3402.50";
                        break;
                    case "16_1": // COUNTRY OF ORIGIN
                        textBoxField.Text = "CA";
                        break;

                    // -------------------------------------------------------------------

                    case "boln": //BOL NUMBER
                        textBoxField.Text = row[datatable.Columns[2]].ToString();
                        break;
                    case "11 LOCAL CARRIER": //CARRIER NAME: FEDEX OR RDWY
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
                    case "class1": //LTL ONLY: First Line of Class
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("fedex"))
                        {
                            textBoxField.Text = "";
                        }
                        else
                        {
                            textBoxField.Text = "55";
                        }
                        break;
                    case "asi1": //CUSTOMER ORDER INFORMATION: ADDITIONAL SHIPPER INFO
                        if (row[datatable.Columns[19]].ToString().ToLower().Equals("fedex"))
                        {
                            textBoxField.Text = "Individual Carton(s)";
                        }
                        else
                        {
                            textBoxField.Text = "Pure Pallet(s)";
                        }
                        break;
                }
            }
        }
    }
    // Export to PDF
    CUSTOM_DOCUMENT.SaveToFile(AppContext.BaseDirectory + "Export/PCB Customs Target " + row[datatable.Columns[1]].ToString() + " " + DateTime.Today.ToString("d") + ".pdf");

}

// Process.Start("explorer.exe", "/select, c:\\teste");
