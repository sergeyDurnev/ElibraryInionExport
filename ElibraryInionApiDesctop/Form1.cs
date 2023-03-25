using ElibraryInionApi.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ElibraryInionApiDesctop
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            comboBox1.Items.Add("Философия");
            comboBox1.Items.Add("История");
            comboBox1.Items.Add("Социология");
            comboBox1.Items.Add("Экономика");
            comboBox1.Items.Add("Правоведение");
            comboBox1.Items.Add("Политология");
            comboBox1.Items.Add("Науковедение");
            comboBox1.Items.Add("Языкознание");
            comboBox1.Items.Add("Литературоведение");
            comboBox1.Items.Add("Религиоведение");

            comboBox2.Enabled = false;
            int yearEnd = DateTime.Now.Year + 1;
            int yearStart = yearEnd - 101;

            while (yearEnd != yearStart)
            {
                comboBox2.Items.Add(yearEnd);
                yearEnd--;
            }

            string codeApiPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
            string apiCode = "";
            using (StreamReader reader = new StreamReader(codeApiPath))
            {
                apiCode = reader.ReadToEnd();
            }
            label3.Text = apiCode;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    richTextBox1.Text = "Получение данных из ELIBRARY";

        //    IElibraryApi elibraryApi = new ElibraryApi();

        //    string outputPath = label5.Text.Replace("|BaseDirectory|", AppDomain.CurrentDomain.BaseDirectory);

        //    if (radioButton2.Checked)
        //    {
        //        string itemIdsData = elibraryApi.GetItemIdsData(false, label3.Text, null, outputPath, textBox1.Text, label8.Text, null, null, textBox2.Text);

        //        XDocument xDoc = XDocument.Parse(itemIdsData);
        //        IEnumerable<XElement> itemIds = xDoc.Root.Descendants("item");
        //        if (itemIds.Count() == 0)
        //        {
        //            richTextBox1.Text += "Ошибка при получении данных по API: " + itemIdsData;
        //        }
        //        else
        //        {
        //            List<Item> items = new List<Item>();
        //            int itemOrder = 0;
        //            foreach (XElement itemIdElement in itemIds)
        //            {
        //                richTextBox1.Text += "\nПолучение данных по адресу: http://elibrary.ru/projects/API-NEB/API_NEB.aspx?ucode="
        //                    + label3.Text + "&sid=26&itemid=" + itemIdElement.Value;
        //                string typeCode = itemIdElement.Attribute("typeCode").Value.Trim();
        //                if (typeCode == "RAR" || typeCode == "REV" || typeCode == "SCO" || typeCode == "BKC" || typeCode == "CLA" || typeCode == "PRC" || typeCode == "THS")
        //                {
        //                    itemOrder++;
        //                    Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
        //                    items.Add(item);
        //                }
        //            }

        //            if (items.Count() > 0)
        //            {
        //                string destinationPath = outputPath;
        //                if (!Directory.Exists(destinationPath))
        //                {
        //                    Directory.CreateDirectory(destinationPath);
        //                }

        //                var groupedItems = items.GroupBy(p => p.ParentId);
        //                foreach (var groupedItem in groupedItems)
        //                {
        //                    string destinationIssuePath = Path.Combine(destinationPath,
        //                         label8.Text + "_" + groupedItem.Key + "_"
        //                        + groupedItem.FirstOrDefault()?.IssueNumber + "_"
        //                        + groupedItem.Count() + ".rtf");
        //                    elibraryApi.GenerateTxtDocument(groupedItem, destinationIssuePath);
        //                }

        //                string dateNow = DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year;

        //                List<ProtocolLog> protocolLogs = new List<ProtocolLog> {new ProtocolLog
        //                {
        //                    Date = dateNow,
        //                    SubjectCode = label8.Text,
        //                    Subject = comboBox1.SelectedItem.ToString(),
        //                    ParentId = items.FirstOrDefault()?.ParentId,
        //                    Number = items.FirstOrDefault()?.IssueNumber,
        //                    Year = items.FirstOrDefault()?.Year,
        //                    RubricCode = textBox2.Text
        //                } };

        //                string itemLogProtocol = "log_protocol_items_" + dateNow + ".xls";
        //                string itemLogProtocolFile = Path.Combine(destinationPath, itemLogProtocol);
        //                elibraryApi.GenerateExcelItemsLog(itemLogProtocolFile, label3.Text, protocolLogs, items);

        //                richTextBox2.Text = $"Год: { items.First().Year} \nНомер: { items.First().IssueNumber } \nВсего статей: {items.Count} " +
        //                    $"\nС аннотациями и ключевыми словами: {items.Where(x => x.Complete == true).Count()} " +
        //                    $"\nБез аннотаций и ключевых слов: {items.Where(x => x.Complete == false).Count()}";

        //                richTextBox1.Text += "\n Выгрузка завершена!";

        //            }
        //        }

        //    }

        //    if (radioButton1.Checked)
        //    {
        //        richTextBox1.Text += "\nПолучение данных по адресу: http://elibrary.ru/projects/API-NEB/API_NEB.aspx?ucode="
        //                   + label3.Text + "&sid=26&itemid=" + textBox1.Text;

        //        Item item = elibraryApi.GetItem(label3.Text, 1, textBox1.Text, label8.Text, textBox1.Text, outputPath, null);
        //        List<Item> items = new List<Item> { item };

        //        string destinationArticlePath = Path.Combine(outputPath,
        //                         label8.Text + "_"
        //                         + textBox1.Text + "_"
        //                        + items.FirstOrDefault()?.IssueNumber + "_"
        //                        + items.Count() + ".rtf");
        //        elibraryApi.GenerateTxtDocument(items, destinationArticlePath);
        //        richTextBox2.Text = $"Год: { items.First().Year} \nНомер: { items.First().IssueNumber } \nВсего статей: {items.Count} " +
        //                    $"\nС аннотациями и ключевыми словами: {items.Where(x => x.Complete == true).Count()} " +
        //                    $"\nБез аннотаций и ключевых слов: {items.Where(x => x.Complete == false).Count()}";

        //        richTextBox1.Text += "\n Выгрузка завершена!";
        //    }

        //    if (radioButton3.Checked)
        //    {
        //        string itemIdsData = elibraryApi.GetItemIdsData(false, label3.Text, null, outputPath, null, label8.Text, textBox1.Text, comboBox2.SelectedItem.ToString(), textBox2.Text);
        //        XDocument xDoc = XDocument.Parse(itemIdsData);
        //        IEnumerable<XElement> itemIds = xDoc.Root.Descendants("item");
        //        if (itemIds.Count() == 0)
        //        {
        //            richTextBox1.Text += "Ошибка при получении данных по API: " + itemIdsData;
        //        }
        //        else
        //        {
        //            List<Item> items = new List<Item>();
        //            int itemOrder = 0;
        //            foreach (XElement itemIdElement in itemIds)
        //            {
        //                richTextBox1.Text += "\nПолучение данных по адресу: http://elibrary.ru/projects/API-NEB/API_NEB.aspx?ucode="
        //                    + label3.Text + "&sid=26&itemid=" + itemIdElement.Value;
        //                string typeCode = itemIdElement.Attribute("typeCode").Value.Trim();
        //                if (typeCode == "RAR" || typeCode == "REV" || typeCode == "SCO" || typeCode == "BKC" || typeCode == "CLA" || typeCode == "PRC" || typeCode == "THS")
        //                {
        //                    itemOrder++;
        //                    Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath);
        //                    items.Add(item);
        //                }
        //            }

        //            if (items.Count() > 0)
        //            {
        //                string dateNow = DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year;
        //                string destinationPath = outputPath;
        //                if (!Directory.Exists(destinationPath))
        //                {
        //                    Directory.CreateDirectory(destinationPath);
        //                }

        //                List<ProtocolLog> protocolLogs = new List<ProtocolLog>();
        //                var groupedItems = items.GroupBy(p => p.ParentId);
        //                foreach (var groupedItem in groupedItems)
        //                {
        //                    string destinationIssuePath = Path.Combine(destinationPath,
        //                         label8.Text + "_" + groupedItem.Key + "_"
        //                        + groupedItem.FirstOrDefault()?.IssueNumber + "_"
        //                        + groupedItem.Count() + ".rtf");
        //                    elibraryApi.GenerateTxtDocument(groupedItem, destinationIssuePath);

        //                    protocolLogs.Add(new ProtocolLog
        //                    {
        //                        Date = dateNow,
        //                        TitleId = textBox1.Text,
        //                        SubjectCode = label8.Text,
        //                        Subject = comboBox1.SelectedItem.ToString(),
        //                        ParentId = groupedItem.Key,
        //                        Number = groupedItem.First().IssueNumber,
        //                        Year = comboBox2.SelectedItem.ToString(),
        //                        RubricCode = textBox2.Text
        //                    });

        //                    richTextBox2.Text += $"Год: { groupedItem.First().Year} \nНомер: { groupedItem.First().IssueNumber } \nВсего статей: {groupedItem.Count()} " +
        //                    $"\nС аннотациями и ключевыми словами: {groupedItem.Where(x => x.Complete == true).Count()} " +
        //                    $"\nБез аннотаций и ключевых слов: {groupedItem.Where(x => x.Complete == false).Count()}\n";
        //                }

        //                string itemLogProtocol = "log_protocol_items_" + dateNow + ".xls";
        //                string itemLogProtocolFile = Path.Combine(destinationPath, itemLogProtocol);
        //                elibraryApi.GenerateExcelItemsLog(itemLogProtocolFile, label3.Text, protocolLogs, items);


        //                richTextBox1.Text += "\n Выгрузка завершена!";

        //            }
        //        }
        //    }
        //}

        //private void label7_Click(object sender, EventArgs e)
        //{

        //}

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedItem.ToString())
            {
                case "Философия":
                    label8.Text = "A02";
                    label8.Visible = true;
                    break;
                case "История":
                    label8.Text = "A03";
                    label8.Visible = true;
                    break;
                case "Социология":
                    label8.Text = "A04";
                    label8.Visible = true;
                    break;
                case "Экономика":
                    label8.Text = "A05";
                    label8.Visible = true;
                    break;
                case "Правоведение":
                    label8.Text = "A10";
                    label8.Visible = true;
                    break;
                case "Политология":
                    label8.Text = "A11";
                    label8.Visible = true;
                    break;
                case "Науковедение":
                    label8.Text = "A12";
                    label8.Visible = true;
                    break;
                case "Языкознание":
                    label8.Text = "A16";
                    label8.Visible = true;
                    break;
                case "Литературоведение":
                    label8.Text = "A17";
                    label8.Visible = true;
                    break;
                case "Религиоведение":
                    label8.Text = "A21";
                    label8.Visible = true;
                    break;
                default:
                    label8.Text = "label8";
                    label8.Visible = false;
                    break;
            }

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                comboBox2.Enabled = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                comboBox2.Enabled = false;
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                comboBox2.Enabled = true;
            }
        }

        private void iconButton1_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "Получение данных из ELIBRARY";

            IElibraryApi elibraryApi = new ElibraryApi();

            string outputPath = label5.Text.Replace("|BaseDirectory|", AppDomain.CurrentDomain.BaseDirectory);

            if (radioButton2.Checked)
            {
                string itemIdsData = elibraryApi.GetItemIdsData(false, label3.Text, null, outputPath, textBox1.Text, label8.Text, null, null, textBox2.Text);

                XDocument xDoc = XDocument.Parse(itemIdsData);
                IEnumerable<XElement> itemIds = xDoc.Root.Descendants("item");
                if (itemIds.Count() == 0)
                {
                    richTextBox1.Text += "Ошибка при получении данных по API: " + itemIdsData;
                }
                else
                {
                    List<Item> items = new List<Item>();
                    int itemOrder = 0;
                    foreach (XElement itemIdElement in itemIds)
                    {
                        richTextBox1.Text += "\nПолучение данных по адресу: http://elibrary.ru/projects/API-NEB/API_NEB.aspx?ucode="
                            + label3.Text + "&sid=26&itemid=" + itemIdElement.Value;
                        string typeCode = itemIdElement.Attribute("typeCode").Value.Trim();

                        if (checkBox1.Checked)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox2.Checked && checkBox2.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox3.Checked && checkBox3.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox4.Checked && checkBox4.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox5.Checked && checkBox5.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox6.Checked && checkBox6.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox7.Checked && checkBox7.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox8.Checked && checkBox8.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox9.Checked && checkBox9.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox10.Checked && checkBox10.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox11.Checked && checkBox11.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox12.Checked && checkBox12.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox13.Checked && checkBox13.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox14.Checked && checkBox14.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                    }

                    if (items.Count() > 0)
                    {
                        string destinationPath = outputPath;
                        if (!Directory.Exists(destinationPath))
                        {
                            Directory.CreateDirectory(destinationPath);
                        }

                        var groupedItems = items.GroupBy(p => p.ParentId);
                        foreach (var groupedItem in groupedItems)
                        {
                            string destinationIssuePath = Path.Combine(destinationPath,
                                 label8.Text + "_" + groupedItem.Key + "_"
                                + groupedItem.FirstOrDefault()?.IssueNumber + "_"
                                + groupedItem.Count() + ".rtf");
                            elibraryApi.GenerateTxtDocument(groupedItem, destinationIssuePath);
                        }

                        string dateNow = DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year;

                        List<ProtocolLog> protocolLogs = new List<ProtocolLog> {new ProtocolLog
                        {
                            Date = dateNow,
                            SubjectCode = label8.Text,
                            Subject = comboBox1.SelectedItem.ToString(),
                            ParentId = items.FirstOrDefault()?.ParentId,
                            Number = items.FirstOrDefault()?.IssueNumber,
                            Year = items.FirstOrDefault()?.Year,
                            RubricCode = textBox2.Text
                        } };

                        string itemLogProtocol = "log_protocol_items_" + dateNow + ".xls";
                        string itemLogProtocolFile = Path.Combine(destinationPath, itemLogProtocol);
                        elibraryApi.GenerateExcelItemsLog(itemLogProtocolFile, label3.Text, protocolLogs, items);

                        richTextBox2.Text = $"Год: {items.First().Year} \nНомер: {items.First().IssueNumber} \nВсего статей: {items.Count} " +
                            $"\nС аннотациями и ключевыми словами: {items.Where(x => x.Complete == true).Count()} " +
                            $"\nБез аннотаций и ключевых слов: {items.Where(x => x.Complete == false).Count()}";

                        richTextBox1.Text += "\n Выгрузка завершена!";

                    }
                }

            }

            if (radioButton1.Checked)
            {
                richTextBox1.Text += "\nПолучение данных по адресу: http://elibrary.ru/projects/API-NEB/API_NEB.aspx?ucode="
                           + label3.Text + "&sid=26&itemid=" + textBox1.Text;

                Item item = elibraryApi.GetItem(label3.Text, 1, textBox1.Text, label8.Text, textBox1.Text, outputPath, null);
                List<Item> items = new List<Item> { item };

                string destinationArticlePath = Path.Combine(outputPath,
                                 label8.Text + "_"
                                 + textBox1.Text + "_"
                                + items.FirstOrDefault()?.IssueNumber + "_"
                                + items.Count() + ".rtf");
                elibraryApi.GenerateTxtDocument(items, destinationArticlePath);
                richTextBox2.Text = $"Год: {items.First().Year} \nНомер: {items.First().IssueNumber} \nВсего статей: {items.Count} " +
                            $"\nС аннотациями и ключевыми словами: {items.Where(x => x.Complete == true).Count()} " +
                            $"\nБез аннотаций и ключевых слов: {items.Where(x => x.Complete == false).Count()}";

                richTextBox1.Text += "\n Выгрузка завершена!";
            }

            if (radioButton3.Checked)
            {
                string itemIdsData = elibraryApi.GetItemIdsData(false, label3.Text, null, outputPath, null, label8.Text, textBox1.Text, comboBox2.SelectedItem.ToString(), textBox2.Text);
                XDocument xDoc = XDocument.Parse(itemIdsData);
                IEnumerable<XElement> itemIds = xDoc.Root.Descendants("item");
                if (itemIds.Count() == 0)
                {
                    richTextBox1.Text += "Ошибка при получении данных по API: " + itemIdsData;
                }
                else
                {
                    List<Item> items = new List<Item>();
                    int itemOrder = 0;
                    foreach (XElement itemIdElement in itemIds)
                    {
                        richTextBox1.Text += "\nПолучение данных по адресу: http://elibrary.ru/projects/API-NEB/API_NEB.aspx?ucode="
                            + label3.Text + "&sid=26&itemid=" + itemIdElement.Value;
                        string typeCode = itemIdElement.Attribute("typeCode").Value.Trim();
                        if (checkBox1.Checked)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox2.Checked && checkBox2.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox3.Checked && checkBox3.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox4.Checked && checkBox4.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox5.Checked && checkBox5.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox6.Checked && checkBox6.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox7.Checked && checkBox7.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox8.Checked && checkBox8.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox9.Checked && checkBox9.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox10.Checked && checkBox10.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox11.Checked && checkBox11.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox12.Checked && checkBox12.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox13.Checked && checkBox13.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                        else if (checkBox14.Checked && checkBox14.Text == typeCode)
                        {
                            itemOrder++;
                            Item item = elibraryApi.GetItem(label3.Text, itemOrder, itemIdElement.Value, label8.Text, textBox1.Text, outputPath, typeCode);
                            items.Add(item);
                        }
                    }

                    if (items.Count() > 0)
                    {
                        string dateNow = DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year;
                        string destinationPath = outputPath;
                        if (!Directory.Exists(destinationPath))
                        {
                            Directory.CreateDirectory(destinationPath);
                        }

                        List<ProtocolLog> protocolLogs = new List<ProtocolLog>();
                        var groupedItems = items.GroupBy(p => p.ParentId);
                        foreach (var groupedItem in groupedItems)
                        {
                            string destinationIssuePath = Path.Combine(destinationPath,
                                 label8.Text + "_" + groupedItem.Key + "_"
                                + groupedItem.FirstOrDefault()?.IssueNumber + "_"
                                + groupedItem.Count() + ".rtf");
                            elibraryApi.GenerateTxtDocument(groupedItem, destinationIssuePath);

                            protocolLogs.Add(new ProtocolLog
                            {
                                Date = dateNow,
                                TitleId = textBox1.Text,
                                SubjectCode = label8.Text,
                                Subject = comboBox1.SelectedItem.ToString(),
                                ParentId = groupedItem.Key,
                                Number = groupedItem.First().IssueNumber,
                                Year = comboBox2.SelectedItem.ToString(),
                                RubricCode = textBox2.Text
                            });

                            richTextBox2.Text += $"Год: {groupedItem.First().Year} \nНомер: {groupedItem.First().IssueNumber} \nВсего статей: {groupedItem.Count()} " +
                            $"\nС аннотациями и ключевыми словами: {groupedItem.Where(x => x.Complete == true).Count()} " +
                            $"\nБез аннотаций и ключевых слов: {groupedItem.Where(x => x.Complete == false).Count()}\n";
                        }

                        string itemLogProtocol = "log_protocol_items_" + dateNow + ".xls";
                        string itemLogProtocolFile = Path.Combine(destinationPath, itemLogProtocol);
                        elibraryApi.GenerateExcelItemsLog(itemLogProtocolFile, label3.Text, protocolLogs, items);


                        richTextBox1.Text += "\n Выгрузка завершена!";

                    }
                }
            }

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_MouseHover(object sender, EventArgs e)
        {
            checkBox2.Text = "аннотация";
        }

        private void checkBox2_MouseLeave(object sender, EventArgs e)
        {
            checkBox2.Text = "ABS";
        }

        private void checkBox3_MouseHover(object sender, EventArgs e)
        {
            checkBox3.Text = "рецензия";
        }

        private void checkBox3_MouseLeave(object sender, EventArgs e)
        {
            checkBox3.Text = "BRV";
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox13.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox14.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox4_MouseHover(object sender, EventArgs e)
        {
            checkBox4.Text = "материалы конференции";
        }

        private void checkBox4_MouseLeave(object sender, EventArgs e)
        {
            checkBox4.Text = "CNF";
        }

        private void checkBox5_MouseHover(object sender, EventArgs e)
        {
            checkBox5.Text = "переписка";
        }

        private void checkBox5_MouseLeave(object sender, EventArgs e)
        {
            checkBox5.Text = "COR";
        }

        private void checkBox6_MouseHover(object sender, EventArgs e)
        {
            checkBox6.Text = "редакторская заметка";
        }

        private void checkBox6_MouseLeave(object sender, EventArgs e)
        {
            checkBox6.Text = "EDI";
        }

        private void checkBox7_MouseHover(object sender, EventArgs e)
        {
            checkBox7.Text = "разное";
        }

        private void checkBox7_MouseLeave(object sender, EventArgs e)
        {
            checkBox7.Text = "MIS";
        }

        private void checkBox8_MouseHover(object sender, EventArgs e)
        {
            checkBox8.Text = "персоналия";
        }

        private void checkBox8_MouseLeave(object sender, EventArgs e)
        {
            checkBox8.Text = "PER";
        }

        private void checkBox9_MouseHover(object sender, EventArgs e)
        {
            checkBox9.Text = "научная статья";
        }

        private void checkBox9_MouseLeave(object sender, EventArgs e)
        {
            checkBox9.Text = "RAR";
        }

        private void checkBox10_MouseHover(object sender, EventArgs e)
        {
            checkBox10.Text = "научный отчет";
        }

        private void checkBox10_MouseLeave(object sender, EventArgs e)
        {
            checkBox10.Text = "REP";
        }

        private void checkBox11_MouseHover(object sender, EventArgs e)
        {
            checkBox11.Text = "обзорная статья";
        }

        private void checkBox11_MouseLeave(object sender, EventArgs e)
        {
            checkBox11.Text = "REV";
        }

        private void checkBox12_MouseHover(object sender, EventArgs e)
        {
            checkBox12.Text = "репринт";
        }

        private void checkBox12_MouseLeave(object sender, EventArgs e)
        {
            checkBox12.Text = "RPR";
        }

        private void checkBox13_MouseHover(object sender, EventArgs e)
        {
            checkBox13.Text = "краткое сообщение";
        }

        private void checkBox13_MouseLeave(object sender, EventArgs e)
        {
            checkBox13.Text = "SCO";
        }

        private void checkBox14_MouseHover(object sender, EventArgs e)
        {
            checkBox14.Text = "не определен";
        }

        private void checkBox14_MouseLeave(object sender, EventArgs e)
        {
            checkBox14.Text = "UNK";
        }

        private void iconButton2_Click(object sender, EventArgs e)
        {
            Form1.ActiveForm.Close();
        }

        // Drag form
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void iconButton3_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void iconButton2_MouseHover(object sender, EventArgs e)
        {
            iconButton2.BackColor = Color.DarkRed;
        }

        private void iconButton2_MouseLeave(object sender, EventArgs e)
        {
            iconButton2.BackColor = Color.DarkOrange;
        }

        private void iconButton3_MouseHover(object sender, EventArgs e)
        {
            iconButton3.BackColor = Color.Orange;
        }

        private void iconButton3_MouseLeave(object sender, EventArgs e)
        {
            iconButton3.BackColor = Color.DarkOrange;
        }
    }
}
