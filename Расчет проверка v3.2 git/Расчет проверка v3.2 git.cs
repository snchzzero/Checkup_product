
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using Word = Microsoft.Office.Interop.Word; // для работы, создания и редактирования документа
using System.IO;
using Microsoft.Office.Interop.Word; // для чтения текстового файла










namespace _2
{
    public partial class Form1 : Form
    {

        const int ImaxPLS = 100;  
        const int ImaxMNS = -100; 
        const double KTMI_Up = 100;  
        const double KTMI_Down = -100; 
        const double DELTA_smech0_mul_Up = 100; 
        const double DELTA_smech0_mul_Down = -100; 

        const double Kdos_Up = 100; 
        const double Kdos_Down = -100; 
        const double KtmDOS_Up = 100; 
        const double KtmDOS_Down = -100; 
        const double Kach_Up = 100; 
        const double Kach_Down = -100; 

        const double TMUdos_sm0_Up = 100; 
        const double TMUdos_sm0_Down = -100; 
        const double Kosc_Up = 100; 
        const double Kosc_Down = -100; 


   
        const double ACH_4Gh_Up = 100; 
        const double ACH_4Gh_Down = -100;
        const double FCH_4Gh_Up = 100; 
        const double FCH_4Gh_Down = -100;
        
        const double ACH_10Gh_Up = 100;  
        const double ACH_10Gh_Down = -100;
        const double FCH_10Gh_Up = 100;  
        const double FCH_10Gh_Down = -100;
       
        const double ACH_40Gh_Up = 100; 
        const double ACH_40Gh_Down = -100;
        const double FCH_40Gh_Up = 100;  
        const double FCH_40Gh_Down = -100;
        
        const double ACH_80Gh_Up = 100; 
        const double ACH_80Gh_Down = -100;
        const double FCH_80Gh_Up = 100; 
        const double FCH_80Gh_Down = -100;
        
        const double ACH_100Gh_Up = 100;  
        const double ACH_100Gh_Down = -100;
        const double FCH_100Gh_Up = 100;  
        const double FCH_100Gh_Down = -100;
     
        const double ACH_400Gh_Up = 100;  
        const double ACH_400Gh_Down = -100;
        



        const double A4Gh_Down = -100;
        const double A4Gh_Up = 100;
        const double B4Gh_Down = -100;
        const double B4Gh_Up = 100;
        
        const double A10Gh_Down = -100;
        const double A10Gh_Up = 100;
        const double B10Gh_Down = -100;
        const double B10Gh_Up = 100;
      
        const double A40Gh_Down = -100;
        const double A40Gh_Up = 100;
        const double B40Gh_Down = -100;
        const double B40Gh_Up = 100;
        
        const double A80Gh_Down = -100;
        const double A80Gh_Up = 100;
        const double B80Gh_Down = -100;
        const double B80Gh_Up = 100;
        
        const double A100Gh_Down = -100;
        const double A100Gh_Up = 100;
        const double B100Gh_Down = -100;
        const double B100Gh_Up = 100;

        //сюда записываем значения при считывание из блокнота для платы 
        double res_A1 = -9999.9;  
        double res_B1 = -9999.9; 
        double res_A2 = -9999.9;  
        double res_B2 = -9999.9;  
        double res_A3 = -9999.9;  
        double res_B3 = -9999.9;  
        double res_A4 = -9999.9;  
        double res_B4 = -9999.9;  
        double res_A5 = -9999.9;  
        double res_B5 = -9999.9;  
        double res_A6 = -9999.9;  
        double res_B6 = -9999.9;  
        double res_A7 = -9999.9; 
        double res_B7 = -9999.9;  

        // записываем значения из блокнота для разных каналов
        // канал 1
        double A4Gh1 = -9999.9;
        double B4Gh1 = -9999.9;
        double A10Gh1 = -9999.9;
        double B10Gh1 = -9999.9;
        double A40Gh1 = -9999.9;
        double B40Gh1 = -9999.9;
        double A80Gh1 = -9999.9;
        double B80Gh1 = -9999.9;
        double A100Gh1 = -9999.9;
        double B100Gh1 = -9999.9;
        // канал 2
        double A4Gh2 = -9999.9;
        double B4Gh2 = -9999.9;
        double A10Gh2 = -9999.9;
        double B10Gh2 = -9999.9;
        double A40Gh2 = -9999.9;
        double B40Gh2 = -9999.9;
        double A80Gh2 = -9999.9;
        double B80Gh2 = -9999.9;
        double A100Gh2 = -9999.9;
        double B100Gh2 = -9999.9;
        // канал 3
        double A4Gh3 = -9999.9;
        double B4Gh3 = -9999.9;
        double A10Gh3 = -9999.9;
        double B10Gh3 = -9999.9;
        double A40Gh3 = -9999.9;
        double B40Gh3 = -9999.9;
        double A80Gh3 = -9999.9;
        double B80Gh3 = -9999.9;
        double A100Gh3 = -9999.9;
        double B100Gh3 = -9999.9;
        // канал 4
        double A4Gh4 = -9999.9;
        double B4Gh4 = -9999.9;
        double A10Gh4 = -9999.9;
        double B10Gh4 = -9999.9;
        double A40Gh4 = -9999.9;
        double B40Gh4 = -9999.9;
        double A80Gh4 = -9999.9;
        double B80Gh4 = -9999.9;
        double A100Gh4 = -9999.9;
        double B100Gh4 = -9999.9;

       // переменные проверки правильности загруженных txt файлов для канала1-4 (РП1-РП4.txt)
        int RP1 = 0;
        int RP2 = 0;
        int RP3 = 0;
        int RP4 = 0;

        // объявление переменных Проверка 1
        double TMIobmPLS; // TM Iобм +, В
        double TMIobmMNS; // TM Iобм -, В
        double ImulPLS; // Iмул +, В
        double ImulMNS; // Iмул -, В
        double TMIobmCM0; // TM Iобм см0, В
        double ImulCM0; // Iмул см0, В
        double IMAXobmPLS; // Imax обм +
        double IMAXobmMNS; // Imax обм -
        double Ktmi; //Ктм i
        double DELmul; //дельта см 0 мульт
        double DELTMIcm0; // дельта ТМ I см 0

        // Проверка 2
        double Kdos;
        double KtmDOS;
        double Kach;

        // Проверка 3
        double TMUdos_sm0; // TM Uдос см0
        double Kosc;

        string operation_0;


        // Переменные для Расчет 
        // записываем значения из блокнота  для разных каналов
        //  №1 (I)
        double IA4Gh1 = -9999.9;
        double IB4Gh1 = -9999.9;
        double IA10Gh1 = -9999.9;
        double IB10Gh1 = -9999.9;
        double IA40Gh1 = -9999.9;
        double IB40Gh1 = -9999.9;
        double IA80Gh1 = -9999.9;
        double IB80Gh1 = -9999.9;
        double IA100Gh1 = -9999.9;
        double IB100Gh1 = -9999.9;
        // канал 2
        double IA4Gh2 = -9999.9;
        double IB4Gh2 = -9999.9;
        double IA10Gh2 = -9999.9;
        double IB10Gh2 = -9999.9;
        double IA40Gh2 = -9999.9;
        double IB40Gh2 = -9999.9;
        double IA80Gh2 = -9999.9;
        double IB80Gh2 = -9999.9;
        double IA100Gh2 = -9999.9;
        double IB100Gh2 = -9999.9;
        // канал 3
        double IA4Gh3 = -9999.9;
        double IB4Gh3 = -9999.9;
        double IA10Gh3 = -9999.9;
        double IB10Gh3 = -9999.9;
        double IA40Gh3 = -9999.9;
        double IB40Gh3 = -9999.9;
        double IA80Gh3 = -9999.9;
        double IB80Gh3 = -9999.9;
        double IA100Gh3 = -9999.9;
        double IB100Gh3 = -9999.9;
        // канал 4
        double IA4Gh4 = -9999.9;
        double IB4Gh4 = -9999.9;
        double IA10Gh4 = -9999.9;
        double IB10Gh4 = -9999.9;
        double IA40Gh4 = -9999.9;
        double IB40Gh4 = -9999.9;
        double IA80Gh4 = -9999.9;
        double IB80Gh4 = -9999.9;
        double IA100Gh4 = -9999.9;
        double IB100Gh4 = -9999.9;
       
        //  (II)
        // канал 1
        double IIA4Gh1 = -9999.9;
        double IIB4Gh1 = -9999.9;
        double IIA10Gh1 = -9999.9;
        double IIB10Gh1 = -9999.9;
        double IIA40Gh1 = -9999.9;
        double IIB40Gh1 = -9999.9;
        double IIA80Gh1 = -9999.9;
        double IIB80Gh1 = -9999.9;
        double IIA100Gh1 = -9999.9;
        double IIB100Gh1 = -9999.9;
        // канал 2
        double IIA4Gh2 = -9999.9;
        double IIB4Gh2 = -9999.9;
        double IIA10Gh2 = -9999.9;
        double IIB10Gh2 = -9999.9;
        double IIA40Gh2 = -9999.9;
        double IIB40Gh2 = -9999.9;
        double IIA80Gh2 = -9999.9;
        double IIB80Gh2 = -9999.9;
        double IIA100Gh2 = -9999.9;
        double IIB100Gh2 = -9999.9;
        // канал 3
        double IIA4Gh3 = -9999.9;
        double IIB4Gh3 = -9999.9;
        double IIA10Gh3 = -9999.9;
        double IIB10Gh3 = -9999.9;
        double IIA40Gh3 = -9999.9;
        double IIB40Gh3 = -9999.9;
        double IIA80Gh3 = -9999.9;
        double IIB80Gh3 = -9999.9;
        double IIA100Gh3 = -9999.9;
        double IIB100Gh3 = -9999.9;
        // канал 4
        double IIA4Gh4 = -9999.9;
        double IIB4Gh4 = -9999.9;
        double IIA10Gh4 = -9999.9;
        double IIB10Gh4 = -9999.9;
        double IIA40Gh4 = -9999.9;
        double IIB40Gh4 = -9999.9;
        double IIA80Gh4 = -9999.9;
        double IIB80Gh4 = -9999.9;
        double IIA100Gh4 = -9999.9;
        double IIB100Gh4 = -9999.9;

        //переменные для записи пути расположения фалов РП1-4 (канал 1-5 ..) *txt для сравнения и невозможности подкрепить один и тот же файл
        // для разных 
        string dIRP1 = "1";  // адрес хранения файла "Проверка .. ВКЗ_РП1" для 1
        string dIRP2 = "2";
        string dIRP3 = "3";
        string dIRP4 = "4";
        string dIIRP1 = "5"; // адрес хранения файла "Проверка ... ВКЗ_РП1" для 2
        string dIIRP2 = "6";
        string dIIRP3 = "7";
        string dIIRP4 = "8";

        // переменные проверки правильности загруженных txt файлов для канала1-4 (РП1-РП4.txt)
        int IRP1 = 0;
        int IIRP1 = 0;
        int IRP2 = 0;
        int IIRP2 = 0;
        int IRP3 = 0;
        int IIRP3 = 0;
        int IRP4 = 0;
        int IIRP4 = 0;

        //private string TemplateFileName = @"C:\Users\evstafyev_aa\Desktop\Prot.docx"; // для офис 2 абсолютный путь к файлу
        private string TemplateFileName = AppDomain.CurrentDomain.BaseDirectory + "/Prot.docx"; // для офис 2 относительный путь к файлу
        private string TemplateFileName2 = AppDomain.CurrentDomain.BaseDirectory + "/Prot2.docx"; // для офис 2 относительный путь к файлу Для протокола  1-4 канал
        private string TemplateFileName3 = AppDomain.CurrentDomain.BaseDirectory + "/Prot3.docx"; // для офис 3 относительный путь к файлу Для протокола  2шт 1-4 канал
        public Form1()
        {
            InitializeComponent();
        }


        private void button7_Click(object sender, EventArgs e) // Кнопка офиса
        {
            if ((checkBox1.Checked == true && checkBox2.Checked == true) || (checkBox1.Checked == false && checkBox2.Checked == false)) //для перед после лакировки
            {
                operation_0 = "не выбрано";
            }

            //проверка 1 вводимые значение
            var nomer = textBox21.Text;
            var operation = operation_0;
            var TMIobmPLS_2 = textBox1.Text;
            var TMIobmMNS_2 = textBox2.Text;
            var ImulPLS_2 = textBox3.Text;
            var ImulMNS_2 = textBox4.Text;
            var TMIobmCM0_2 = textBox5.Text;
            var ImulCM0_2 = textBox6.Text;

            //проверка 1 расчитанные значение
            var IMAXobmPLS_2 = textBox7.Text;
            var IMAXobmMNS_2 = textBox8.Text;
            var Ktmi_2 = textBox9.Text;
            var DELmul_2 = textBox10.Text;
            var DELTMIcm0_2 = textBox11.Text;

            //проверка 2 вводимые значения
            var TMUdosPLS_2 = textBox19.Text;
            var TMUdosMNS_2 = textBox18.Text;
            var UsumPLS_2 = textBox17.Text;
            var UsumMNS_2 = textBox16.Text;
            var TMIobmK_PLS_2 = textBox15.Text;
            var TMIobmK_MNS_2 = textBox14.Text;

            //проверка 2 расчитанные значение
            var Kdos_2 = textBox13.Text;
            var KtmDOS_2 = textBox12.Text;
            var Kach_2 = textBox20.Text;

            //проверка 3
            var TMUdos_sm0_2 = textBox29.Text;
            var TMIosc_KPLS_2 = textBox27.Text;
            var TMIosc_KMNS_2 = textBox26.Text;
            var Kosc_2 = textBox23.Text;

            //проверка 4 - ...
            //var res_A1_word = textBox25.Text;
            //var res_A2_word = textBox30.Text;
            //var res_A3_word = textBox32.Text;
            //var res_A4_word = textBox34.Text;
            //var res_A5_word = textBox35.Text;
            //var res_A6_word = textBox36.Text;
            //var res_A7_word = textBox37.Text;
            //var res_B1_word = textBox28.Text;
            //var res_B2_word = textBox31.Text;
            //var res_B3_word = textBox33.Text;
            //var res_B4_word = textBox38.Text;
            //var res_B5_word = textBox39.Text;
            //var res_B6_word = textBox40.Text;
            //var res_B7_word = textBox41.Text;

            string DATA_s = System.DateTime.Now.ToShortDateString(); //дата сокращенная
            string Time_s = System.DateTime.Now.ToShortTimeString(); //время короткое
            var Family = textBox22.Text; // Фамилия И.О. проверяющего
            string DATA_ss = System.DateTime.Now.Year.ToString() + "." + System.DateTime.Now.Month.ToString("d2") + "." + System.DateTime.Now.Day.ToString();


            var wordApp = new Word.Application();//создаем переменную "wordApp" с приложением оболочки ворда
            wordApp.Visible = false; //не видеть в процессе экспорта открытое окно ворда

            try
            {
                // работа с шаблоном
                var wordDocument = wordApp.Documents.Open(TemplateFileName);//открываем документ


                // меняем форматирование текста в заисимости от полученных результатов (подсвечиваем красным если вне диапозона) ред. 2021.04.07
                // взято с ресурса https://fooobar.com/questions/516118/c-searching-a-text-in-word-and-getting-the-range-of-the-result
                Word.Range range;
                Word.Range temprange;
                Word.Selection currentSelection;
                // для Проверки-1 
                // IMAXobmPLS Imax обм +
                if (0 > IMAXobmPLS || IMAXobmPLS > ImaxPLS)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{7az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{7az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // IMAXobmMNS Imax обм -
                if (ImaxMNS > IMAXobmMNS || IMAXobmMNS > 0)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{8az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{8az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                //  Ktmi //Ктм i
                if (KTMI_Down > Ktmi || Ktmi > KTMI_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{9az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{9az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // DELmul дельта см 0 мульт
                if (DELTA_smech0_mul_Down > DELmul || DELmul > DELTA_smech0_mul_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{10az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{10az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // DELTMIcm0  дельта ТМ I см 0
                if (DELTA_smech0_mul_Down > DELTMIcm0 || DELTMIcm0 > DELTA_smech0_mul_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{11az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{11az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // Для Проверки-2
                // Kdos
                if (Kdos_Down > Kdos || Kdos > Kdos_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{18az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{18az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // KtmDOS
                if (KtmDOS_Down > KtmDOS || KtmDOS > KtmDOS_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{19az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{19az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // Kach
                if (Kach_Down > Kach || Kach > Kach_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{20az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{20az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // TMUdos_sm0
                if (TMUdos_sm0_Down > TMUdos_sm0 || TMUdos_sm0 > TMUdos_sm0_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{21az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{21az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // Kosc
                if (Kosc_Down > Kosc || Kosc > Kosc_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{24az}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{24az}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // для АЧХ
                // {4Gh_a}
                if (ACH_4Gh_Down > res_A2 || res_A2 > ACH_4Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{4Gh_a}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{4Gh_a}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {10Gh_a}
                if (ACH_10Gh_Down > res_A3 || res_A3 > ACH_10Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{10Gh_a}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{10Gh_a}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {40Gh_a}
                if (ACH_40Gh_Down > res_A4 || res_A4 > ACH_40Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{40Gh_a}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{40Gh_a}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {80Gh_a}
                if (ACH_80Gh_Down > res_A5 || res_A5 > ACH_80Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{80Gh_a}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{80Gh_a}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {100Gh_a}
                if (ACH_100Gh_Down > res_A6 || res_A6 > ACH_100Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{100Gh_a}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{100Gh_a}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {400Gh_a}
                if (ACH_400Gh_Down > res_A7 || res_A7 > ACH_400Gh_Up)
                {
                    wordApp.Selection.Find.Execute("{400Gh_a}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{400Gh_a}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // Для ...
                // {4Gh_b}
                if (FCH_4Gh_Down > res_B2 || res_B2 > FCH_4Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{4Gh_b}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{4Gh_b}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {10Gh_b}
                if (FCH_10Gh_Down > res_B3 || res_B3 > FCH_10Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{10Gh_b}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{10Gh_b}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {40Gh_b}
                if (FCH_40Gh_Down > res_B4 || res_B4 > FCH_40Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{40Gh_b}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{40Gh_b}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {80Gh_b}
                if (FCH_80Gh_Down > res_B5 || res_B5 > FCH_80Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{80Gh_b}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{80Gh_b}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {100Gh_b}
                if (FCH_100Gh_Down > res_B6 || res_B6 > FCH_100Gh_Up)
                {
                    wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp.Selection.Find.Execute("{100Gh_b}");
                    range = wordApp.Selection.Range;
                    if (range.Text.Contains("{100Gh_b}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                // {400Gh_b} не предъявляется







                // пример правильного кода по поиску текста и замене форматирования
                //Word.Range range;
                //Word.Range temprange;
                //Word.Selection currentSelection;
                //wordApp.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; // применяем данную строчку кода если ранне уже выполнялся поиск и выделение
                //wordApp.Selection.Find.Execute("{Family}");
                //range = wordApp.Selection.Range;
                //if (range.Text.Contains("{Family}"))
                //{
                //    // gets desired range here it gets last character to make superscript in range
                //    temprange = wordDocument.Range(range.Start, range.End);
                //    temprange.Select();
                //    currentSelection = wordApp.Selection;
                //    currentSelection.Font.Color = Word.WdColor.wdColorRed;
                //}




                //поиск по документу и замена
                ReplaceWord_Ofice2("{operation}", operation_0, wordDocument);
                ReplaceWord_Ofice2("{nomer}", nomer, wordDocument);
                ReplaceWord_Ofice2("{1az}", TMIobmPLS_2, wordDocument);
                ReplaceWord_Ofice2("{2az}", TMIobmMNS_2, wordDocument);
                ReplaceWord_Ofice2("{3az}", ImulPLS_2, wordDocument);
                ReplaceWord_Ofice2("{4az}", ImulMNS_2, wordDocument);
                ReplaceWord_Ofice2("{5az}", TMIobmCM0_2, wordDocument);
                ReplaceWord_Ofice2("{6az}", ImulCM0_2, wordDocument);

                ReplaceWord_Ofice2("{7az}", IMAXobmPLS_2, wordDocument);
                ReplaceWord_Ofice2("{8az}", IMAXobmMNS_2, wordDocument);
                ReplaceWord_Ofice2("{9az}", Ktmi_2, wordDocument);
                ReplaceWord_Ofice2("{10az}", DELmul_2, wordDocument);
                ReplaceWord_Ofice2("{11az}", DELTMIcm0_2, wordDocument);

                ReplaceWord_Ofice2("{12az}", TMUdosPLS_2, wordDocument);
                ReplaceWord_Ofice2("{13az}", TMUdosMNS_2, wordDocument);
                ReplaceWord_Ofice2("{14az}", UsumPLS_2, wordDocument);
                ReplaceWord_Ofice2("{15az}", UsumMNS_2, wordDocument);
                ReplaceWord_Ofice2("{16az}", TMIobmK_PLS_2, wordDocument);
                ReplaceWord_Ofice2("{17az}", TMIobmK_MNS_2, wordDocument);

                ReplaceWord_Ofice2("{18az}", Kdos_2, wordDocument);
                ReplaceWord_Ofice2("{19az}", KtmDOS_2, wordDocument);
                ReplaceWord_Ofice2("{20az}", Kach_2, wordDocument);

                ReplaceWord_Ofice2("{21az}", TMUdos_sm0_2, wordDocument);
                ReplaceWord_Ofice2("{22az}", TMIosc_KPLS_2, wordDocument);
                ReplaceWord_Ofice2("{23az}", TMIosc_KMNS_2, wordDocument);
                ReplaceWord_Ofice2("{24az}", Kosc_2, wordDocument);

                //для ...
                if (res_A1 != -9999.9 && res_B7 != -9999.9 && res_A5 != -9999.9)
                {
                    ReplaceWord_Ofice2("{1Gh_a}", res_A1.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{1Gh_b}", res_B1.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{4Gh_a}", res_A2.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{4Gh_b}", res_B2.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{10Gh_a}", res_A3.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{10Gh_b}", res_B3.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{40Gh_a}", res_A4.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{40Gh_b}", res_B4.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{80Gh_a}", res_A5.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{80Gh_b}", res_B5.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{100Gh_a}", res_A6.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{100Gh_b}", res_B6.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{400Gh_a}", res_A7.ToString("f2"), wordDocument);
                    ReplaceWord_Ofice2("{400Gh_b}", res_B7.ToString("f2"), wordDocument);
                }
                if (res_A1 == -9999.9 && res_B7 == -9999.9 && res_A5 == -9999.9)
                {
                    ReplaceWord_Ofice2("{1Gh_a}", "", wordDocument);
                    ReplaceWord_Ofice2("{1Gh_b}", "", wordDocument);
                    ReplaceWord_Ofice2("{4Gh_a}", "", wordDocument);
                    ReplaceWord_Ofice2("{4Gh_b}", "", wordDocument);
                    ReplaceWord_Ofice2("{10Gh_a}", "", wordDocument);
                    ReplaceWord_Ofice2("{10Gh_b}", "", wordDocument);
                    ReplaceWord_Ofice2("{40Gh_a}", "", wordDocument);
                    ReplaceWord_Ofice2("{40Gh_b}", "", wordDocument);
                    ReplaceWord_Ofice2("{80Gh_a}", "", wordDocument);
                    ReplaceWord_Ofice2("{80Gh_b}", "", wordDocument);
                    ReplaceWord_Ofice2("{100Gh_a}", "", wordDocument);
                    ReplaceWord_Ofice2("{100Gh_b}", "", wordDocument);
                    ReplaceWord_Ofice2("{400Gh_a}", "", wordDocument);
                    ReplaceWord_Ofice2("{400Gh_b}", "", wordDocument);
                }
                ReplaceWord_Ofice2("{Family}", Family, wordDocument);
                ReplaceWord_Ofice2("{data}", DATA_s, wordDocument);
                ReplaceWord_Ofice2("{time}", Time_s, wordDocument);



                wordDocument.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "PCB №" + nomer + "_" + DATA_ss + ".docx");
                //wordDocument.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + DATA_ss + "Prot_№ " + nomer + ".docx");
                wordApp.Visible = true;




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }
        private void ReplaceWord_Ofice2(string stubToReplace, string text, Word.Document wordDocument)// для офис 2 поиск и замена в документе
        {
            var range = wordDocument.Content; // диапозон символов в документе внутри всего текстового документа "Content" 
            range.Find.ClearFormatting(); //сброс поиска во всем документе
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);    // выполнить поиск 
        }


        //Проверка 1.
        void Indikator_Repultat_Prov_1(double IMAXobmPLS, double IMAXobmMNS, double Ktmi, double DELmul, double DELTMIcm0)// Окрашивание формулы в зеленый если верное значение и в красный если неверное значение
        {
            // Максимальный ток в нагрузке
            if ((IMAXobmPLS <= ImaxPLS) && (IMAXobmPLS > 0))
            {
                label9.ForeColor = System.Drawing.Color.Green;
                label11.ForeColor = System.Drawing.Color.Green;
                label10.ForeColor = System.Drawing.Color.Green;
            }
            if (IMAXobmPLS > ImaxPLS || IMAXobmPLS < 0)
            {
                label9.ForeColor = System.Drawing.Color.Red;
                label11.ForeColor = System.Drawing.Color.Red;
                label10.ForeColor = System.Drawing.Color.Red;
            }
            if (IMAXobmMNS >= ImaxMNS && IMAXobmMNS < 0)
            {
                label16.ForeColor = System.Drawing.Color.Green;
                label15.ForeColor = System.Drawing.Color.Green;
                label14.ForeColor = System.Drawing.Color.Green;
            }
            if (IMAXobmMNS < ImaxMNS || IMAXobmMNS > 0)
            {
                label16.ForeColor = System.Drawing.Color.Red;
                label15.ForeColor = System.Drawing.Color.Red;
                label14.ForeColor = System.Drawing.Color.Red;
            }

            // Коэффициент предачи по цепи телеметрии тока
            if (Ktmi >= KTMI_Down && Ktmi <= KTMI_Up)
            {
                label12.ForeColor = System.Drawing.Color.Green;
                label13.ForeColor = System.Drawing.Color.Green;
            }
            if (Ktmi < KTMI_Down || Ktmi > KTMI_Up)
            {
                label12.ForeColor = System.Drawing.Color.Red;
                label13.ForeColor = System.Drawing.Color.Red;
            }

            // Проверка нулевого значения тока в цепи нагрузки УМ при отсутствии упр.сигналов на входах в мА 
            if (DELmul >= DELTA_smech0_mul_Down && DELmul <= DELTA_smech0_mul_Up)
            {
                label19.ForeColor = System.Drawing.Color.Green;
                label20.ForeColor = System.Drawing.Color.Green;
                label21.ForeColor = System.Drawing.Color.Green;
            }
            if (DELmul < DELTA_smech0_mul_Down || DELmul > DELTA_smech0_mul_Up)
            {
                label19.ForeColor = System.Drawing.Color.Red;
                label20.ForeColor = System.Drawing.Color.Red;
                label21.ForeColor = System.Drawing.Color.Red;
            }
            // Проверка нулевого значения тока в цепи нагрузки УМ при отсутствии упр.сигналов на входах в мА для TM I
            if (DELTMIcm0 >= DELTA_smech0_mul_Down && DELTMIcm0 <= DELTA_smech0_mul_Up)
            {
                label22.ForeColor = System.Drawing.Color.Green;
                label23.ForeColor = System.Drawing.Color.Green;
                label24.ForeColor = System.Drawing.Color.Green;
            }
            if (DELTMIcm0 < DELTA_smech0_mul_Down || DELTMIcm0 > DELTA_smech0_mul_Up)
            {
                label22.ForeColor = System.Drawing.Color.Red;
                label23.ForeColor = System.Drawing.Color.Red;
                label24.ForeColor = System.Drawing.Color.Red;
            }

        }
        void OutPrint(double IMAXobmPLS, double IMAXobmMNS, double Ktmi, double DELmul, double DELTMIcm0) // Проверка 1. Вывод результата на экран
        {
            textBox7.Text = IMAXobmPLS.ToString("f2");
            textBox8.Text = IMAXobmMNS.ToString("f2");
            textBox9.Text = Ktmi.ToString("f3");
            textBox10.Text = DELmul.ToString("f3");
            textBox11.Text = DELTMIcm0.ToString("f3");

        }
        void ClearOut() // Для проверки 1. если сброс
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            label9.ForeColor = System.Drawing.Color.Black;
            label11.ForeColor = System.Drawing.Color.Black;
            label10.ForeColor = System.Drawing.Color.Black;
            label16.ForeColor = System.Drawing.Color.Black;
            label15.ForeColor = System.Drawing.Color.Black;
            label14.ForeColor = System.Drawing.Color.Black;
            label12.ForeColor = System.Drawing.Color.Black;
            label13.ForeColor = System.Drawing.Color.Black;
            label19.ForeColor = System.Drawing.Color.Black;
            label20.ForeColor = System.Drawing.Color.Black;
            label21.ForeColor = System.Drawing.Color.Black;
            label22.ForeColor = System.Drawing.Color.Black;
            label23.ForeColor = System.Drawing.Color.Black;
            label24.ForeColor = System.Drawing.Color.Black;
            label69.Text = "";
            label70.Text = "";
            label71.Text = "";
            label72.Text = "";
            label73.Text = "";

        }
        private void button1_Click(object sender, EventArgs e) //Проверка 1. Кнопка "Расчет"
        {
            //откуда считываются переменные? из окна ввода
            TMIobmPLS = Convert.ToDouble(textBox1.Text);
            TMIobmMNS = Convert.ToDouble(textBox2.Text);
            ImulPLS = Convert.ToDouble(textBox3.Text);
            ImulMNS = Convert.ToDouble(textBox4.Text);
            TMIobmCM0 = Convert.ToDouble(textBox5.Text);
            ImulCM0 = Convert.ToDouble(textBox6.Text);

            // расчет Imax Обм + и Imax Обм -
            IMAXobmPLS = (ImulPLS * 1000) / 130;
            IMAXobmMNS = (ImulMNS * 1000) / 130;

            // расчет Ктм i
            Ktmi = (TMIobmPLS - TMIobmMNS) / (IMAXobmPLS - IMAXobmMNS);

            //расчет дельта мул см0 
            DELmul = (ImulCM0 * 1000) / 130;

            //расчет дельта ТМ I см0 
            DELTMIcm0 = TMIobmCM0 / Ktmi;

            //под результатами расчета показываем диапозон допустимых значений
            label69.Text = "(0..." + ImaxPLS + ")";
            label70.Text = "(" + ImaxMNS + "...0)";
            label71.Text = "(" + KTMI_Down + "..." + KTMI_Up + ")";
            label72.Text = "(" + DELTA_smech0_mul_Down + "..." + DELTA_smech0_mul_Up + ")";
            label73.Text = "(" + DELTA_smech0_mul_Down + "..." + DELTA_smech0_mul_Up + ")";

            //окрашивание формулы в зеленый если верное значение и в красный если неверное значение
            Indikator_Repultat_Prov_1(IMAXobmPLS, IMAXobmMNS, Ktmi, DELmul, DELTMIcm0);

            OutPrint(IMAXobmPLS, IMAXobmMNS, Ktmi, DELmul, DELTMIcm0);

            // после нажатия кнопки расчет убираем подпись "введите значение"
            if (textBox1.Text != "")
            {
                label103.Text = "";
            }
            if (textBox2.Text != "")
            {
                label110.Text = "";
            }
            if (textBox3.Text != "")
            {
                label117.Text = "";
            }
            if (textBox4.Text != "")
            {
                label118.Text = "";
            }
            if (textBox5.Text != "")
            {
                label119.Text = "";
            }
            if (textBox6.Text != "")
            {
                label120.Text = "";
            }

        }
        private void button2_Click(object sender, EventArgs e)// Проверка 1. Кнопка "Сброс"
        {
            //если нажимаем кнопку Сброс, то очищаются поля
            ClearOut();
            // после нажатия кнопки сброс добавляем подпись "введите значение"
            label103.Text = "(введите значение)";
            label110.Text = "(введите значение)";
            label117.Text = "(введите значение)";
            label118.Text = "(введите значение)";
            label119.Text = "(введите значение)";
            label120.Text = "(введите значение)";
        }

        //Проверка 2. 
        void Indikator_Repultat_Prov_2(double Kdos, double KtmDOS, double Kach, double UsumPLS, double UsumMNS) //Окрашивание формулы в зеленый если верное значение и в красный если неверное значение
        {
            //Значение статического коэффициента передачи по цепи "СигнДОС-вход суммирующего ОУ"
            if (Kdos <= Kdos_Up && Kdos >= Kdos_Down)
            {
                label30.ForeColor = System.Drawing.Color.Green;
                label29.ForeColor = System.Drawing.Color.Green;
            }
            if (Kdos > Kdos_Up || Kdos < Kdos_Down)
            {
                label30.ForeColor = System.Drawing.Color.Red;
                label29.ForeColor = System.Drawing.Color.Red;
            }
            //Значение статического коэффициента по епи телеметрии ДОС
            if (KtmDOS <= KtmDOS_Up && KtmDOS >= KtmDOS_Down)
            {
                label25.ForeColor = System.Drawing.Color.Green;
                label27.ForeColor = System.Drawing.Color.Green;
            }
            if (KtmDOS > KtmDOS_Up || KtmDOS < KtmDOS_Down)
            {
                label25.ForeColor = System.Drawing.Color.Red;
                label27.ForeColor = System.Drawing.Color.Red;
            }
            //Статический коэффициент передачи по цепи СигнДОС-ток в нагрузке УМ
            if (Kach <= Kach_Up && Kach >= Kach_Down)
            {
                label62.ForeColor = System.Drawing.Color.Green;
                label63.ForeColor = System.Drawing.Color.Green;
            }
            if (Kach > Kach_Up || Kach < Kach_Down)
            {
                label62.ForeColor = System.Drawing.Color.Red;
                label63.ForeColor = System.Drawing.Color.Red;
            }

            //По каналу DA4-1k значение должно быть отрицательным (UsumPLS)
            if (UsumPLS < 0)
            {
                label58.ForeColor = System.Drawing.Color.Green;
                label34.ForeColor = System.Drawing.Color.Green;
                label35.ForeColor = System.Drawing.Color.Green;
            }
            if (UsumPLS > 0)
            {
                label58.ForeColor = System.Drawing.Color.Red;
                label34.ForeColor = System.Drawing.Color.Red;
                label35.ForeColor = System.Drawing.Color.Red;
            }

            //По каналу DA4-1k значение должно быть положительным (UsumPLS)
            if (UsumMNS > 0)
            {
                label59.ForeColor = System.Drawing.Color.Green;
                label60.ForeColor = System.Drawing.Color.Green;
                label61.ForeColor = System.Drawing.Color.Green;
            }
            if (UsumMNS < 0)
            {
                label59.ForeColor = System.Drawing.Color.Red;
                label60.ForeColor = System.Drawing.Color.Red;
                label61.ForeColor = System.Drawing.Color.Red;
            }







        }
        void OutPrint2(double Kdos, double KtmDOS, double Kach) //Проверка 2. Вывод результата на экран
        {
            textBox13.Text = Kdos.ToString("f3");
            textBox12.Text = KtmDOS.ToString("f3");
            textBox20.Text = Kach.ToString("f3");

        }
        void ClearOut2() // Проверка 2. если сброс
        {
            textBox19.Text = "";
            textBox18.Text = "";
            textBox17.Text = "";
            textBox16.Text = "";
            textBox15.Text = "";
            textBox14.Text = "";
            textBox13.Text = "";
            textBox12.Text = "";
            textBox20.Text = "";
            label29.ForeColor = System.Drawing.Color.Black;
            label30.ForeColor = System.Drawing.Color.Black;
            label25.ForeColor = System.Drawing.Color.Black;
            label27.ForeColor = System.Drawing.Color.Black;
            label62.ForeColor = System.Drawing.Color.Black;
            label63.ForeColor = System.Drawing.Color.Black;
            label80.Text = "";
            label81.Text = "";
            label82.Text = "";
            label34.ForeColor = System.Drawing.Color.Black;
            label35.ForeColor = System.Drawing.Color.Black;
            label58.ForeColor = System.Drawing.Color.Black;
            label59.ForeColor = System.Drawing.Color.Black;
            label60.ForeColor = System.Drawing.Color.Black;
            label61.ForeColor = System.Drawing.Color.Black;

        }
        private void button4_Click(object sender, EventArgs e) // Проверка 2. Кнопка "Расчет"
        {
            // объявление переменных
            double TMUdosPLS; // TM Uдос +
            double TMUdosMNS; // TM Uдос -
            double UsumPLS; // Uсум+
            double UsumMNS; // Uсум-
            double TMIobmK_PLS; // TM Iobm K+
            double TMIobmK_MNS; // TM Iobm K-
            double Ktmi2;

            //откуда считываются переменные? из окна ввода
            TMUdosPLS = Convert.ToDouble(textBox19.Text);
            TMUdosMNS = Convert.ToDouble(textBox18.Text);
            UsumPLS = Convert.ToDouble(textBox17.Text);
            UsumMNS = Convert.ToDouble(textBox16.Text);
            TMIobmK_PLS = Convert.ToDouble(textBox15.Text);
            TMIobmK_MNS = Convert.ToDouble(textBox14.Text);
            Ktmi2 = Convert.ToDouble(textBox9.Text);
            //расчет Кдос         
            Kdos = Math.Abs(UsumPLS - UsumMNS) / 2;

            //расчет Kтм дос
            KtmDOS = (TMUdosPLS - TMUdosMNS) / Math.Abs(UsumPLS - UsumMNS);

            //Расчет Kэч
            Kach = (TMIobmK_PLS - TMIobmK_MNS) / (2 * Ktmi2);

            //под результатами расчета показываем диапозон допустимых значений 
            label80.Text = "(" + Kdos_Down + "..." + Kdos_Up + ")";
            label81.Text = "(" + KtmDOS_Down + "..." + KtmDOS_Up + ")";
            label82.Text = "(" + Kach_Down + "..." + Kach_Up + ")";


            Indikator_Repultat_Prov_2(Kdos, KtmDOS, Kach, UsumPLS, UsumMNS);
            //настройка вывода результата
            OutPrint2(Kdos, KtmDOS, Kach);

            // после нажатия кнопки расчет убираем подпись "введите значение"
            if (textBox19.Text != "")
            {
                label121.Text = "";
            }
            if (textBox18.Text != "")
            {
                label122.Text = "";
            }
            if (textBox17.Text != "")
            {
                label123.Text = "";
            }
            if (textBox16.Text != "")
            {
                label124.Text = "";
            }
            if (textBox15.Text != "")
            {
                label125.Text = "";
            }
            if (textBox14.Text != "")
            {
                label126.Text = "";
            }
        }
        private void button3_Click(object sender, EventArgs e) // Проверка 2. Кнопка "Сброс
        {
            // При нажатии на кнопку очищаются поля
            ClearOut2();

            // после нажатия кнопки сброс добавляем подпись "введите значение"
            label121.Text = "(введите значение)";
            label122.Text = "(введите значение)";
            label123.Text = "(введите значение)";
            label124.Text = "(введите значение)";
            label125.Text = "(введите значение)";
            label126.Text = "(введите значение)";
        }

        //Проверка 3. 
        void Indikator_Repultat_Prov_3(double TMUdos_sm0, double Kosc, double TMIosc_KPLS, double TMIosc_KMNS)//Окрашивание формулы в зеленый если верное значение и в красный если неверное значение
        {
            //TM Uдос см0 показания телеметрии
            if (TMUdos_sm0 >= TMUdos_sm0_Down && TMUdos_sm0 <= TMUdos_sm0_Up)
            {
                label83.ForeColor = System.Drawing.Color.Green;
                label84.ForeColor = System.Drawing.Color.Green;
                label94.ForeColor = System.Drawing.Color.Green;
            }
            if (TMUdos_sm0 < TMUdos_sm0_Down || TMUdos_sm0 > TMUdos_sm0_Up)
            {
                label83.ForeColor = System.Drawing.Color.Red;
                label84.ForeColor = System.Drawing.Color.Red;
                label94.ForeColor = System.Drawing.Color.Red;
            }

            //проверка стат.коэф Косц
            if (Kosc <= Kosc_Up && Kosc >= Kosc_Down)
            {
                label90.ForeColor = System.Drawing.Color.Green;
                label91.ForeColor = System.Drawing.Color.Green;
            }
            if (Kosc > Kosc_Up || Kosc < Kosc_Down)
            {
                label90.ForeColor = System.Drawing.Color.Red;
                label91.ForeColor = System.Drawing.Color.Red;
            }

            //проверка TMIosc_KPLS,  TMIosc_KMNS положительное значение или отрицательное
            if (TMIosc_KPLS < 0)
            {
                label77.ForeColor = System.Drawing.Color.Green;
                label78.ForeColor = System.Drawing.Color.Green;
                label79.ForeColor = System.Drawing.Color.Green;
            }
            if (TMIosc_KPLS > 0)
            {
                label77.ForeColor = System.Drawing.Color.Red;
                label78.ForeColor = System.Drawing.Color.Red;
                label79.ForeColor = System.Drawing.Color.Red;
            }
            if (TMIosc_KMNS > 0)
            {
                label74.ForeColor = System.Drawing.Color.Green;
                label75.ForeColor = System.Drawing.Color.Green;
                label76.ForeColor = System.Drawing.Color.Green;
            }
            if (TMIosc_KMNS < 0)
            {
                label74.ForeColor = System.Drawing.Color.Red;
                label75.ForeColor = System.Drawing.Color.Red;
                label76.ForeColor = System.Drawing.Color.Red;
            }


        }
        void OutPrint3(double Kosc) //Проверка 3. Вывод результата на экран
        {
            textBox23.Text = Kosc.ToString("f3");
        }
        void ClearOut3() // Проверка 3. если сброс
        {
            textBox23.Text = "";
            textBox26.Text = "";
            textBox27.Text = "";
            textBox29.Text = "";
            label83.ForeColor = System.Drawing.Color.Black;
            label84.ForeColor = System.Drawing.Color.Black;
            label94.ForeColor = System.Drawing.Color.Black;
            label90.ForeColor = System.Drawing.Color.Black;
            label91.ForeColor = System.Drawing.Color.Black;
            label85.Text = "";
            label86.Text = "";
            label74.ForeColor = System.Drawing.Color.Black;
            label75.ForeColor = System.Drawing.Color.Black;
            label76.ForeColor = System.Drawing.Color.Black;
            label77.ForeColor = System.Drawing.Color.Black;
            label78.ForeColor = System.Drawing.Color.Black;
            label79.ForeColor = System.Drawing.Color.Black;

            // после нажатия кнопки сброс добавляем подпись "введите значение"
            label127.Text = "(введите значение)";
            label128.Text = "(введите значение)";
            label129.Text = "(введите значение)";
        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        private void button6_Click(object sender, EventArgs e) //Проверка 3. Кнопка "Расчет"
        {

            double TMIosc_KPLS; //ТМ I осц к+
            double TMIosc_KMNS; //ТМ I осц к-
            double Ktmi3; // Kтм i

            TMUdos_sm0 = Convert.ToDouble(textBox29.Text);
            TMIosc_KPLS = Convert.ToDouble(textBox27.Text);
            TMIosc_KMNS = Convert.ToDouble(textBox26.Text);
            Ktmi3 = Convert.ToDouble(textBox9.Text);

            //расчет Косц  
            Kosc = Math.Abs(TMIosc_KPLS - TMIosc_KMNS) / Ktmi3;

            //под результатами расчета показываем диапозон допустимых значений 
            label85.Text = "(" + TMUdos_sm0_Down + "..." + TMUdos_sm0_Up + ")";
            label86.Text = "(" + Kosc_Down + "..." + Kosc_Up + ")";

            Indikator_Repultat_Prov_3(TMUdos_sm0, Kosc, TMIosc_KPLS, TMIosc_KMNS);
            OutPrint3(Kosc);

            // после нажатия кнопки расчет убираем подпись "введите значение"
            if (textBox29.Text != "")
            {
                label127.Text = "";
            }
            if (textBox27.Text != "")
            {
                label128.Text = "";
            }
            if (textBox26.Text != "")
            {
                label129.Text = "";
            }


        }
        private void button5_Click(object sender, EventArgs e) // Проверка 3. Кнопка "Сброс
        {
            ClearOut3();
        }


        private void textBox1_KeyPress(object sender, KeyPressEventArgs e) // Если в окно ввода данных "textBox" запрет на ввод точки "." 
        {
            if (e.KeyChar == '.') e.KeyChar = ',';
            if (e.KeyChar == '.') e.Handled = true;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            // checkBox1 = (CheckBox)sender;
            if (checkBox1.Checked == true)
            {
                operation_0 = "";
                operation_0 = "перед";
            }

        }



        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox = (CheckBox)sender;
            if (checkBox2.Checked == true)
            {
                operation_0 = "";
                operation_0 = "после";
            }


        }

        OpenFileDialog ofd = new OpenFileDialog();

        // Проверка 4.... для платы
        private void button8_Click(object sender, EventArgs e) // Проверка 4 ..., кнопка открыть файл
        {
            //tb24
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox24.Text = ofd.FileName;
            }

            try
            {
                string path = textBox24.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {

                    string line;
                    string a;
                    int total = 0;




                    while ((line = sr.ReadLine()) != null)
                    {
                        a = line;
                        if (total == 1) // поиск нужных нам цифр в строке 1 --- 1 
                        {
                            string[] res = a.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            res_A1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            res_B1 = Convert.ToDouble(res[3]);
                            textBox25.Text = res_A1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox28.Text = res_B1.ToString("f4");
                        }
                        if (total == 2) // поиск нужных нам цифр в строке 2 --- 4 
                        {
                            string[] res = a.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            res_A2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            res_B2 = Convert.ToDouble(res[3]);
                            textBox30.Text = res_A2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox31.Text = res_B2.ToString("f4");
                        }
                        if (total == 3) // поиск нужных нам цифр в строке 3 --- 10 
                        {
                            string[] res = a.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            res_A3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            res_B3 = Convert.ToDouble(res[3]);
                            textBox32.Text = res_A3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox33.Text = res_B3.ToString("f4");
                        }
                        if (total == 4) // поиск нужных нам цифр в строке 3 --- 40 
                        {
                            string[] res = a.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            res_A4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            res_B4 = Convert.ToDouble(res[3]);
                            textBox34.Text = res_A4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox38.Text = res_B4.ToString("f4");
                        }
                        if (total == 5) // поиск нужных нам цифр в строке 3 --- 80 
                        {
                            string[] res = a.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            res_A5 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            res_B5 = Convert.ToDouble(res[3]);
                            textBox35.Text = res_A5.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox39.Text = res_B5.ToString("f4");
                        }
                        if (total == 6) // поиск нужных нам цифр в строке 3 --- 100 
                        {
                            string[] res = a.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            res_A6 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            res_B6 = Convert.ToDouble(res[3]);
                            textBox36.Text = res_A6.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox40.Text = res_B6.ToString("f4");
                        }
                        if (total == 7) // поиск нужных нам цифр в строке 3 --- 400
                        {
                            string[] res = a.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            res_A7 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            res_B7 = Convert.ToDouble(res[3]);
                            textBox37.Text = res_A7.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox41.Text = res_B7.ToString("f4");
                        }



                        ++total;
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }

        private void button10_Click(object sender, EventArgs e)  // Проверка 4 -  при нажатии на кнопку расчет:
        {
            // показываем диапозон значений согласно ТУ
            label104.Text = "(" + ACH_4Gh_Down + "..." + ACH_4Gh_Up + ")";
            label105.Text = "(" + ACH_10Gh_Down + "..." + ACH_10Gh_Up + ")";
            label106.Text = "(" + ACH_40Gh_Down + "..." + ACH_40Gh_Up + ")";
            label107.Text = "(" + ACH_80Gh_Down + "..." + ACH_80Gh_Up + ")";
            label108.Text = "(" + ACH_100Gh_Down + "..." + ACH_100Gh_Up + ")";
            label109.Text = "(" + ACH_400Gh_Down + "..." + ACH_400Gh_Up + ")";

            label111.Text = "(" + FCH_4Gh_Down + "..." + FCH_4Gh_Up + ")";
            label112.Text = "(" + FCH_10Gh_Down + "..." + FCH_10Gh_Up + ")";
            label113.Text = "(" + FCH_40Gh_Down + "..." + FCH_40Gh_Up + ")";
            label114.Text = "(" + FCH_80Gh_Down + "..." + FCH_80Gh_Up + ")";
            label115.Text = "(" + FCH_100Gh_Down + "..." + FCH_100Gh_Up + ")";
            label116.Text = "( не предъявляется )";  // для ... не предъявляется значение

            if (res_A2 != -9999.9 && res_B5 != -9999.9) // проверяем введено ли значение в текст бокс
            {
                // производим перерасчет из значений частот .... вычитаем значение на ... и записываем результат в текст бокс
                // ..
                res_A2 = res_A2 - res_A1;  // ...
                textBox30.Text = res_A2.ToString("f2");
                if (ACH_4Gh_Down <= res_A2 && res_A2 <= ACH_4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox30.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox30.BackColor = System.Drawing.Color.Red;
                }

                res_B2 = res_B2 - res_B1;  // ...
                textBox31.Text = res_B2.ToString("f4");
                if (FCH_4Gh_Down <= res_B2 && res_B2 <= FCH_4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox31.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox31.BackColor = System.Drawing.Color.Red;
                }

                // ...
                res_A3 = res_A3 - res_A1;  // ...
                textBox32.Text = res_A3.ToString("f2");
                if (ACH_10Gh_Down <= res_A3 && res_A3 <= ACH_10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox32.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox32.BackColor = System.Drawing.Color.Red;
                }

                res_B3 = res_B3 - res_B1;  // ...
                textBox33.Text = res_B3.ToString("f4");
                if (FCH_10Gh_Down <= res_B3 && res_B3 <= FCH_10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox33.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox33.BackColor = System.Drawing.Color.Red;
                }

                // ...
                res_A4 = res_A4 - res_A1;  // ....
                textBox34.Text = res_A4.ToString("f2");
                if (ACH_40Gh_Down <= res_A4 && res_A4 <= ACH_40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox34.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox34.BackColor = System.Drawing.Color.Red;
                }

                res_B4 = res_B4 - res_B1;  // ....
                textBox38.Text = res_B4.ToString("f4");
                if (FCH_40Gh_Down <= res_B4 && res_B4 <= FCH_40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox38.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox38.BackColor = System.Drawing.Color.Red;
                }


                // ...
                res_A5 = res_A5 - res_A1;  // ....
                textBox35.Text = res_A5.ToString("f2");
                if (ACH_80Gh_Down <= res_A5 && res_A5 <= ACH_80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox35.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox35.BackColor = System.Drawing.Color.Red;
                }

                res_B5 = res_B5 - res_B1;  // ....
                textBox39.Text = res_B5.ToString("f4");
                if (FCH_80Gh_Down <= res_B5 && res_B5 <= FCH_80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox39.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox39.BackColor = System.Drawing.Color.Red;
                }


                // ....
                res_A6 = res_A6 - res_A1;  // ....
                textBox36.Text = res_A6.ToString("f2");
                if (ACH_100Gh_Down <= res_A6 && res_A6 <= ACH_100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox36.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox36.BackColor = System.Drawing.Color.Red;
                }

                res_B6 = res_B6 - res_B1;  // ....
                textBox40.Text = res_B6.ToString("f4");
                if (FCH_100Gh_Down <= res_B6 && res_B6 <= FCH_100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox40.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox40.BackColor = System.Drawing.Color.Red;
                }


                // ...
                res_A7 = res_A7 - res_A1;  // ....
                textBox37.Text = res_A7.ToString("f2");
                if (ACH_400Gh_Down <= res_A7 && res_A7 <= ACH_400Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox37.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox37.BackColor = System.Drawing.Color.Red;
                }

                res_B7 = res_B7 - res_B1;  // ...
                textBox41.Text = res_B7.ToString("f4");
                // не предъявляется
            }
        }

        private void button9_Click(object sender, EventArgs e) // Проверка 4 - ... Если нажать кнопку сброс
        {
            textBox24.Text = "";
            textBox25.Text = "";
            textBox25.BackColor = System.Drawing.Color.White;
            textBox28.Text = "";
            textBox28.BackColor = System.Drawing.Color.White;
            textBox30.Text = "";
            textBox30.BackColor = System.Drawing.Color.White;
            textBox31.Text = "";
            textBox31.BackColor = System.Drawing.Color.White;
            textBox32.Text = "";
            textBox32.BackColor = System.Drawing.Color.White;
            textBox33.Text = "";
            textBox33.BackColor = System.Drawing.Color.White;
            textBox34.Text = "";
            textBox34.BackColor = System.Drawing.Color.White;
            textBox38.Text = "";
            textBox38.BackColor = System.Drawing.Color.White;
            textBox35.Text = "";
            textBox35.BackColor = System.Drawing.Color.White;
            textBox39.Text = "";
            textBox39.BackColor = System.Drawing.Color.White;
            textBox36.Text = "";
            textBox36.BackColor = System.Drawing.Color.White;
            textBox40.Text = "";
            textBox40.BackColor = System.Drawing.Color.White;
            textBox37.Text = "";
            textBox37.BackColor = System.Drawing.Color.White;
            textBox41.Text = "";
            textBox41.BackColor = System.Drawing.Color.White;
            label104.Text = "";
            label105.Text = "";
            label106.Text = "";
            label107.Text = "";
            label108.Text = "";
            label109.Text = "";
            label111.Text = "";
            label112.Text = "";
            label113.Text = "";
            label114.Text = "";
            label115.Text = "";
            label116.Text = "";
            res_A1 = -9999.9;  // ..
            res_B1 = -9999.9;  // ..
            res_A2 = -9999.9;  // ...
            res_B2 = -9999.9;  // ..
            res_A3 = -9999.9;  // ...
            res_B3 = -9999.9;  // ..
            res_A4 = -9999.9;  // ...
            res_B4 = -9999.9;  // ...
            res_A5 = -9999.9;  // ...
            res_B5 = -9999.9;  //..
            res_A6 = -9999.9;  // ...
            res_B6 = -9999.9;  // ...
            res_A7 = -9999.9;  // ...
            res_B7 = -9999.9;  // ...
        }


        // Проверка ... Для протокола версии 1 где проверяется... в кол-ве 1 шт
        private void button11_Click(object sender, EventArgs e) // Кнопка открыть РП1 (канал 1 ..) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox82.Text = ofd.FileName;
                string prov_ofd = textBox82.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП1.txt")
                {
                    textBox82.BackColor = System.Drawing.Color.LightGreen;
                    label162.Text = "";
                    RP1 = 0;

                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП1.txt")
                {
                    textBox82.BackColor = System.Drawing.Color.Red;
                    label162.Text = "(выбран некорректный файл)";
                    RP1 = 1;
                }
            }

            try
            {
                string path = textBox82.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A4Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B4Gh1 = Convert.ToDouble(res[3]);
                            textBox43.Text = A4Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox49.Text = B4Gh1.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A10Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B10Gh1 = Convert.ToDouble(res[3]);
                            textBox44.Text = A10Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox50.Text = B10Gh1.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A40Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B40Gh1 = Convert.ToDouble(res[3]);
                            textBox45.Text = A40Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox51.Text = B40Gh1.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A80Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B80Gh1 = Convert.ToDouble(res[3]);
                            textBox47.Text = A80Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox52.Text = B80Gh1.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A100Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B100Gh1 = Convert.ToDouble(res[3]);
                            textBox48.Text = A100Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox53.Text = B100Gh1.ToString("f2");
                        }





                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e) // Кнопка открыть РП2 (канал 2 ...) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox83.Text = ofd.FileName;
                string prov_ofd = textBox83.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП2.txt")
                {
                    textBox83.BackColor = System.Drawing.Color.LightGreen;
                    label163.Text = "";
                    RP2 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП2.txt")
                {
                    textBox83.BackColor = System.Drawing.Color.Red;
                    label163.Text = "(выбран некорректный файл)";
                    RP2 = 1;
                }
            }

            try
            {
                string path = textBox83.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A4Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B4Gh2 = Convert.ToDouble(res[3]);
                            textBox42.Text = A4Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox57.Text = B4Gh2.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A10Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B10Gh2 = Convert.ToDouble(res[3]);
                            textBox46.Text = A10Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox58.Text = B10Gh2.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A40Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B40Gh2 = Convert.ToDouble(res[3]);
                            textBox54.Text = A40Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox59.Text = B40Gh2.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A80Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B80Gh2 = Convert.ToDouble(res[3]);
                            textBox55.Text = A80Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox60.Text = B80Gh2.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A100Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B100Gh2 = Convert.ToDouble(res[3]);
                            textBox56.Text = A100Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox61.Text = B100Gh2.ToString("f2");
                        }





                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e) // Кнопка открыть РП3 (канал 3 .) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox84.Text = ofd.FileName;
                string prov_ofd = textBox84.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП3.txt")
                {
                    textBox84.BackColor = System.Drawing.Color.LightGreen;
                    label164.Text = "";
                    RP3 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП3.txt")
                {
                    textBox84.BackColor = System.Drawing.Color.Red;
                    label164.Text = "(выбран некорректный файл)";
                    RP3 = 1;
                }
            }


            try
            {
                string path = textBox84.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A4Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B4Gh3 = Convert.ToDouble(res[3]);
                            textBox62.Text = A4Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox67.Text = B4Gh3.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A10Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B10Gh3 = Convert.ToDouble(res[3]);
                            textBox63.Text = A10Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox68.Text = B10Gh3.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A40Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B40Gh3 = Convert.ToDouble(res[3]);
                            textBox64.Text = A40Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox69.Text = B40Gh3.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A80Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B80Gh3 = Convert.ToDouble(res[3]);
                            textBox65.Text = A80Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox70.Text = B80Gh3.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A100Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B100Gh3 = Convert.ToDouble(res[3]);
                            textBox66.Text = A100Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox71.Text = B100Gh3.ToString("f2");
                        }





                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button14_Click(object sender, EventArgs e) // Кнопка открыть РП4 (канал 4 ..) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox85.Text = ofd.FileName;
                string prov_ofd = textBox85.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП4.txt")
                {
                    textBox85.BackColor = System.Drawing.Color.LightGreen;
                    label165.Text = "";
                    RP4 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП4.txt")
                {
                    textBox85.BackColor = System.Drawing.Color.Red;
                    label165.Text = "(выбран некорректный файл)";
                    RP4 = 1;
                }
            }

            try
            {
                string path = textBox85.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A4Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B4Gh4 = Convert.ToDouble(res[3]);
                            textBox72.Text = A4Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox77.Text = B4Gh4.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A10Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B10Gh4 = Convert.ToDouble(res[3]);
                            textBox73.Text = A10Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox78.Text = B10Gh4.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A40Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B40Gh4 = Convert.ToDouble(res[3]);
                            textBox74.Text = A40Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox79.Text = B40Gh4.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A80Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B80Gh4 = Convert.ToDouble(res[3]);
                            textBox75.Text = A80Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox80.Text = B80Gh4.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            A100Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            B100Gh4 = Convert.ToDouble(res[3]);
                            textBox76.Text = A100Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox81.Text = B100Gh4.ToString("f2");
                        }





                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button16_Click(object sender, EventArgs e) // Кнопка проверка - проверяем значения .. с каналов 1 - 4 в соответствии с требованиями ТУ
        {
            label148.Text = "Требования ТУ";
            label151.Text = "#";
            label150.Text = "#";
            label156.Text = A4Gh_Down + "..." + A4Gh_Up;
            label161.Text = B4Gh_Down + "..." + B4Gh_Up;
            label155.Text = A10Gh_Down + "..." + A10Gh_Up;
            label160.Text = B10Gh_Down + "..." + B10Gh_Up;
            label154.Text = A40Gh_Down + "..." + A40Gh_Up;
            label159.Text = B40Gh_Down + "..." + B40Gh_Up;
            label153.Text = A80Gh_Down + "..." + A80Gh_Up;
            label158.Text = B80Gh_Down + "..." + B80Gh_Up;
            label152.Text = A100Gh_Down + "..." + A100Gh_Up;
            label157.Text = B100Gh_Down + "..." + B100Gh_Up;

            // Окрашиваем текст бокс в зеленый или красный в зависимости от вхождения в диапозон
            // ....
            if (A4Gh_Down <= A4Gh1 && A4Gh1 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox43.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox43.BackColor = System.Drawing.Color.Red;
            }
            if (A4Gh_Down <= A4Gh2 && A4Gh2 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox42.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox42.BackColor = System.Drawing.Color.Red;
            }
            if (A4Gh_Down <= A4Gh3 && A4Gh3 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox62.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox62.BackColor = System.Drawing.Color.Red;
            }
            if (A4Gh_Down <= A4Gh4 && A4Gh4 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox72.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox72.BackColor = System.Drawing.Color.Red;
            }

            // ...
            if (A10Gh_Down <= A10Gh1 && A10Gh1 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox44.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox44.BackColor = System.Drawing.Color.Red;
            }
            if (A10Gh_Down <= A10Gh2 && A10Gh2 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox46.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox46.BackColor = System.Drawing.Color.Red;
            }
            if (A10Gh_Down <= A10Gh3 && A10Gh3 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox63.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox63.BackColor = System.Drawing.Color.Red;
            }
            if (A10Gh_Down <= A10Gh4 && A10Gh4 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox73.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox73.BackColor = System.Drawing.Color.Red;
            }

            // ...
            if (A40Gh_Down <= A40Gh1 && A40Gh1 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox45.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox45.BackColor = System.Drawing.Color.Red;
            }
            if (A40Gh_Down <= A40Gh2 && A40Gh2 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox54.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox54.BackColor = System.Drawing.Color.Red;
            }
            if (A40Gh_Down <= A40Gh3 && A40Gh3 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox64.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox64.BackColor = System.Drawing.Color.Red;
            }
            if (A40Gh_Down <= A40Gh4 && A40Gh4 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox74.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox74.BackColor = System.Drawing.Color.Red;
            }

            // ...
            if (A80Gh_Down <= A80Gh1 && A80Gh1 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox47.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox47.BackColor = System.Drawing.Color.Red;
            }
            if (A80Gh_Down <= A80Gh2 && A80Gh2 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox55.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox55.BackColor = System.Drawing.Color.Red;
            }
            if (A80Gh_Down <= A80Gh3 && A80Gh3 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox65.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox65.BackColor = System.Drawing.Color.Red;
            }
            if (A80Gh_Down <= A80Gh4 && A80Gh4 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox75.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox75.BackColor = System.Drawing.Color.Red;
            }

            //....
            if (A100Gh_Down <= A100Gh1 && A100Gh1 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox48.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox48.BackColor = System.Drawing.Color.Red;
            }
            if (A100Gh_Down <= A100Gh2 && A100Gh2 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox56.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox56.BackColor = System.Drawing.Color.Red;
            }
            if (A100Gh_Down <= A100Gh3 && A100Gh3 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox66.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox66.BackColor = System.Drawing.Color.Red;
            }
            if (A100Gh_Down <= A100Gh4 && A100Gh4 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox76.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox76.BackColor = System.Drawing.Color.Red;
            }


            // ....
            if (B4Gh_Down <= B4Gh1 && B4Gh1 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox49.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox49.BackColor = System.Drawing.Color.Red;
            }
            if (B4Gh_Down <= B4Gh2 && B4Gh2 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox57.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox57.BackColor = System.Drawing.Color.Red;
            }
            if (B4Gh_Down <= B4Gh3 && B4Gh3 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox67.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox67.BackColor = System.Drawing.Color.Red;
            }
            if (B4Gh_Down <= B4Gh4 && B4Gh4 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox77.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox77.BackColor = System.Drawing.Color.Red;
            }

            // ...
            if (B10Gh_Down <= B10Gh1 && B10Gh1 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox50.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox50.BackColor = System.Drawing.Color.Red;
            }
            if (B10Gh_Down <= B10Gh2 && B10Gh2 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox58.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox58.BackColor = System.Drawing.Color.Red;
            }
            if (B10Gh_Down <= B10Gh3 && B10Gh3 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox68.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox68.BackColor = System.Drawing.Color.Red;
            }
            if (B10Gh_Down <= B10Gh4 && B10Gh4 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox78.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox78.BackColor = System.Drawing.Color.Red;
            }

            // ...
            if (B40Gh_Down <= B40Gh1 && B40Gh1 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox51.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox51.BackColor = System.Drawing.Color.Red;
            }
            if (B40Gh_Down <= B40Gh2 && B40Gh2 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox59.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox59.BackColor = System.Drawing.Color.Red;
            }
            if (B40Gh_Down <= B40Gh3 && B40Gh3 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox69.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox69.BackColor = System.Drawing.Color.Red;
            }
            if (B40Gh_Down <= B40Gh4 && B40Gh4 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox79.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox79.BackColor = System.Drawing.Color.Red;
            }

            // ....
            if (B80Gh_Down <= B80Gh1 && B80Gh1 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox52.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox52.BackColor = System.Drawing.Color.Red;
            }
            if (B80Gh_Down <= B80Gh2 && B80Gh2 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox60.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox60.BackColor = System.Drawing.Color.Red;
            }
            if (B80Gh_Down <= B80Gh3 && B80Gh3 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox70.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox70.BackColor = System.Drawing.Color.Red;
            }
            if (B80Gh_Down <= B80Gh4 && B80Gh4 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox80.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox80.BackColor = System.Drawing.Color.Red;
            }

            // ....
            if (B100Gh_Down <= B100Gh1 && B100Gh1 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox53.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox53.BackColor = System.Drawing.Color.Red;
            }
            if (B100Gh_Down <= B100Gh2 && B100Gh2 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox61.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox61.BackColor = System.Drawing.Color.Red;
            }
            if (B100Gh_Down <= B100Gh3 && B100Gh3 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox71.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox71.BackColor = System.Drawing.Color.Red;
            }
            if (B100Gh_Down <= B100Gh4 && B100Gh4 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox81.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox81.BackColor = System.Drawing.Color.Red;
            }
        }

        private void button17_Click(object sender, EventArgs e) // Если нажимаем кнопку сброс
        {
            label148.Text = "";
            label151.Text = "";
            label150.Text = "";
            label156.Text = "";
            label161.Text = "";
            label155.Text = "";
            label160.Text = "";
            label154.Text = "";
            label159.Text = "";
            label153.Text = "";
            label158.Text = "";
            label152.Text = "";
            label157.Text = "";

            label162.Text = "";
            label163.Text = "";
            label164.Text = "";
            label165.Text = "";

            textBox82.Text = "";
            textBox82.BackColor = System.Drawing.Color.White;
            textBox83.Text = "";
            textBox83.BackColor = System.Drawing.Color.White;
            textBox84.Text = "";
            textBox84.BackColor = System.Drawing.Color.White;
            textBox85.Text = "";
            textBox85.BackColor = System.Drawing.Color.White;
            // 1 канал
            textBox43.Text = "";
            textBox43.BackColor = System.Drawing.Color.White;
            textBox44.Text = "";
            textBox44.BackColor = System.Drawing.Color.White;
            textBox45.Text = "";
            textBox45.BackColor = System.Drawing.Color.White;
            textBox47.Text = "";
            textBox47.BackColor = System.Drawing.Color.White;
            textBox48.Text = "";
            textBox48.BackColor = System.Drawing.Color.White;
            textBox49.Text = "";
            textBox49.BackColor = System.Drawing.Color.White;
            textBox50.Text = "";
            textBox50.BackColor = System.Drawing.Color.White;
            textBox51.Text = "";
            textBox51.BackColor = System.Drawing.Color.White;
            textBox52.Text = "";
            textBox52.BackColor = System.Drawing.Color.White;
            textBox53.Text = "";
            textBox53.BackColor = System.Drawing.Color.White;
            // 2 канал
            textBox42.Text = "";
            textBox42.BackColor = System.Drawing.Color.White;
            textBox46.Text = "";
            textBox46.BackColor = System.Drawing.Color.White;
            textBox54.Text = "";
            textBox54.BackColor = System.Drawing.Color.White;
            textBox55.Text = "";
            textBox55.BackColor = System.Drawing.Color.White;
            textBox56.Text = "";
            textBox56.BackColor = System.Drawing.Color.White;
            textBox57.Text = "";
            textBox57.BackColor = System.Drawing.Color.White;
            textBox58.Text = "";
            textBox58.BackColor = System.Drawing.Color.White;
            textBox59.Text = "";
            textBox59.BackColor = System.Drawing.Color.White;
            textBox60.Text = "";
            textBox60.BackColor = System.Drawing.Color.White;
            textBox61.Text = "";
            textBox61.BackColor = System.Drawing.Color.White;
            // 3 канал
            textBox62.Text = "";
            textBox62.BackColor = System.Drawing.Color.White;
            textBox63.Text = "";
            textBox63.BackColor = System.Drawing.Color.White;
            textBox64.Text = "";
            textBox64.BackColor = System.Drawing.Color.White;
            textBox65.Text = "";
            textBox65.BackColor = System.Drawing.Color.White;
            textBox66.Text = "";
            textBox66.BackColor = System.Drawing.Color.White;
            textBox67.Text = "";
            textBox67.BackColor = System.Drawing.Color.White;
            textBox68.Text = "";
            textBox68.BackColor = System.Drawing.Color.White;
            textBox69.Text = "";
            textBox69.BackColor = System.Drawing.Color.White;
            textBox70.Text = "";
            textBox70.BackColor = System.Drawing.Color.White;
            textBox71.Text = "";
            textBox71.BackColor = System.Drawing.Color.White;
            // 4 канал
            textBox72.Text = "";
            textBox72.BackColor = System.Drawing.Color.White;
            textBox73.Text = "";
            textBox73.BackColor = System.Drawing.Color.White;
            textBox74.Text = "";
            textBox74.BackColor = System.Drawing.Color.White;
            textBox75.Text = "";
            textBox75.BackColor = System.Drawing.Color.White;
            textBox76.Text = "";
            textBox76.BackColor = System.Drawing.Color.White;
            textBox77.Text = "";
            textBox77.BackColor = System.Drawing.Color.White;
            textBox78.Text = "";
            textBox78.BackColor = System.Drawing.Color.White;
            textBox79.Text = "";
            textBox79.BackColor = System.Drawing.Color.White;
            textBox80.Text = "";
            textBox80.BackColor = System.Drawing.Color.White;
            textBox81.Text = "";
            textBox81.BackColor = System.Drawing.Color.White;

            textBox86.Text = ""; // номер

            // канал 1
            A4Gh1 = -9999.9;
            B4Gh1 = -9999.9;
            A10Gh1 = -9999.9;
            B10Gh1 = -9999.9;
            A40Gh1 = -9999.9;
            B40Gh1 = -9999.9;
            A80Gh1 = -9999.9;
            B80Gh1 = -9999.9;
            A100Gh1 = -9999.9;
            B100Gh1 = -9999.9;
            // канал 2
            A4Gh2 = -9999.9;
            B4Gh2 = -9999.9;
            A10Gh2 = -9999.9;
            B10Gh2 = -9999.9;
            A40Gh2 = -9999.9;
            B40Gh2 = -9999.9;
            A80Gh2 = -9999.9;
            B80Gh2 = -9999.9;
            A100Gh2 = -9999.9;
            B100Gh2 = -9999.9;
            // канал 3
            A4Gh3 = -9999.9;
            B4Gh3 = -9999.9;
            A10Gh3 = -9999.9;
            B10Gh3 = -9999.9;
            A40Gh3 = -9999.9;
            B40Gh3 = -9999.9;
            A80Gh3 = -9999.9;
            B80Gh3 = -9999.9;
            A100Gh3 = -9999.9;
            B100Gh3 = -9999.9;
            // канал 4
            A4Gh4 = -9999.9;
            B4Gh4 = -9999.9;
            A10Gh4 = -9999.9;
            B10Gh4 = -9999.9;
            A40Gh4 = -9999.9;
            B40Gh4 = -9999.9;
            A80Gh4 = -9999.9;
            B80Gh4 = -9999.9;
            A100Gh4 = -9999.9;
            B100Gh4 = -9999.9;
        }

        private void button15_Click(object sender, EventArgs e) // Вывести протокол проверки ... в Microsoft Word 
        {
            var wordApp2 = new Word.Application();//создаем переменную "wordApp2" с приложением оболочки ворда
            wordApp2.Visible = false; //не видеть в процессе экспорта открытое окно ворда

            try
            {
                var nomer = textBox86.Text; // номер ..
                var Family = textBox22.Text; // Фамилия И.О. проверяющего  
                string DATA_s = System.DateTime.Now.ToShortDateString(); //дата сокращенная
                string Time_s = System.DateTime.Now.ToShortTimeString(); //время короткое
                string DATA_ss = System.DateTime.Now.Year.ToString() + "." + System.DateTime.Now.Month.ToString("d2") + "." + System.DateTime.Now.Day.ToString();
                // работа с шаблоном
                var wordDocument2 = wordApp2.Documents.Open(TemplateFileName2);//открываем документ


                // меняем форматирование текста в заисимости от полученных результатов (подсвечиваем красным если вне диапозона) ред. 2021.04.07
                // взято с ресурса https://fooobar.com/questions/516118/c-searching-a-text-in-word-and-getting-the-range-of-the-result
                Word.Range range;
                Word.Range temprange;
                Word.Selection currentSelection;
                // ...
                if (A4Gh_Down > A4Gh1 || A4Gh1 > A4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A4Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A4Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > A4Gh2 || A4Gh2 > A4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A4Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A4Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > A4Gh3 || A4Gh3 > A4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A4Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A4Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > A4Gh4 || A4Gh4 > A4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A4Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A4Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ...
                if (A10Gh_Down > A10Gh1 || A10Gh1 > A10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A10Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A10Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > A10Gh2 || A10Gh2 > A10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A10Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A10Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > A10Gh3 || A10Gh3 > A10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A10Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A10Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > A10Gh4 || A10Gh4 > A10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A10Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A10Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ....
                if (A40Gh_Down > A40Gh1 || A40Gh1 > A40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A40Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A40Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > A40Gh2 || A40Gh2 > A40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A40Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A40Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > A40Gh3 || A40Gh3 > A40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A40Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A40Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > A40Gh4 || A40Gh4 > A40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A40Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A40Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ...
                if (A80Gh_Down > A80Gh1 || A80Gh1 > A80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A80Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A80Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > A80Gh2 || A80Gh2 > A80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A80Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A80Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > A80Gh3 || A80Gh3 > A80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A80Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A80Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > A80Gh4 || A80Gh4 > A80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A80Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A80Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ....
                if (A100Gh_Down > A100Gh1 || A100Gh1 > A100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A100Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A100Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > A100Gh2 || A100Gh2 > A100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A100Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A100Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > A100Gh3 || A100Gh3 > A100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A100Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A100Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > A100Gh4 || A100Gh4 > A100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{A100Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{A100Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }


                // ...
                if (B4Gh_Down > B4Gh1 || B4Gh1 > B4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B4Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B4Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > B4Gh2 || B4Gh2 > B4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B4Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B4Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > B4Gh3 || B4Gh3 > B4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B4Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B4Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > B4Gh4 || B4Gh4 > B4Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B4Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B4Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ...
                if (B10Gh_Down > B10Gh1 || B10Gh1 > B10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B10Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B10Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > B10Gh2 || B10Gh2 > B10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B10Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B10Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > B10Gh3 || B10Gh3 > B10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B10Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B10Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > B10Gh4 || B10Gh4 > B10Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B10Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B10Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ...
                if (B40Gh_Down > B40Gh1 || B40Gh1 > B40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B40Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B40Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > B40Gh2 || B40Gh2 > B40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B40Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B40Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > B40Gh3 || B40Gh3 > B40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B40Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B40Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > B40Gh4 || B40Gh4 > B40Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B40Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B40Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ...
                if (B80Gh_Down > B80Gh1 || B80Gh1 > B80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B80Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B80Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > B80Gh2 || B80Gh2 > B80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B80Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B80Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > B80Gh3 || B80Gh3 > B80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B80Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B80Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > B80Gh4 || B80Gh4 > B80Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B80Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B80Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // ...
                if (B100Gh_Down > B100Gh1 || B100Gh1 > B100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B100Gh1}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B100Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > B100Gh2 || B100Gh2 > B100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B100Gh2}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B100Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > B100Gh3 || B100Gh3 > B100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B100Gh3}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B100Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > B100Gh4 || B100Gh4 > B100Gh_Up)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("{B100Gh4}");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("{B100Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }


                // проверки правильности загруженных txt файлов для канала1-4 (РП1-РП4.txt)
                // Если расчитан ... для канала1 не из файла "Проверка ... ВКЗ_РП1"  то в протоколе канал 1 подсвечиваеся красным и тд для остальных каналов
                if (RP1 == 1)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("канал 1");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("канал 1"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (RP2 == 1)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("канал 2");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("канал 2"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (RP3 == 1)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("канал 3");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("канал 3"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (RP4 == 1)
                {
                    wordApp2.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp2.Selection.Find.Execute("канал 4");
                    range = wordApp2.Selection.Range;
                    if (range.Text.Contains("канал 4"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument2.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp2.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // поиск и замена для...
                if (A4Gh1 == -9999.9 && B4Gh1 == -9999.9 && B100Gh4 == -9999.9) // если файл не выбран и нет значений
                {
                    ReplaceWord_Ofice2("{A4Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A4Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A4Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A4Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh4}", "", wordDocument2);
                    // Для ...2 поиск и замена для ...
                    ReplaceWord_Ofice2("{B4Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B4Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B4Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B4Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh4}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh1}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh2}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh3}", "", wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh4}", "", wordDocument2);
                }
                if (A4Gh1 != -9999.9 && B4Gh1 != -9999.9 && B100Gh4 != -9999.9) // если файл выбран и подставлены значения в текст бокс
                {
                    ReplaceWord_Ofice2("{A4Gh1}", A4Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A4Gh2}", A4Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A4Gh3}", A4Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A4Gh4}", A4Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh1}", A10Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh2}", A10Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh3}", A10Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A10Gh4}", A10Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh1}", A40Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh2}", A40Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh3}", A40Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A40Gh4}", A40Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh1}", A80Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh2}", A80Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh3}", A80Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A80Gh4}", A80Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh1}", A100Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh2}", A100Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh3}", A100Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{A100Gh4}", A100Gh4.ToString(), wordDocument2);
                    // поиск и замена для...
                    ReplaceWord_Ofice2("{B4Gh1}", B4Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B4Gh2}", B4Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B4Gh3}", B4Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B4Gh4}", B4Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh1}", B10Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh2}", B10Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh3}", B10Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B10Gh4}", B10Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh1}", B40Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh2}", B40Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh3}", B40Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B40Gh4}", B40Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh1}", B80Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh2}", B80Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh3}", B80Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B80Gh4}", B80Gh4.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh1}", B100Gh1.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh2}", B100Gh2.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh3}", B100Gh3.ToString(), wordDocument2);
                    ReplaceWord_Ofice2("{B100Gh4}", B100Gh4.ToString(), wordDocument2);
                }

                if (checkBox3.Checked == true)
                {
                    ReplaceWord_Ofice2("{ИЗД№0000}", "", wordDocument2);
                    ReplaceWord_Ofice2("{family}", "", wordDocument2);
                    ReplaceWord_Ofice2("{date}", "", wordDocument2);
                    ReplaceWord_Ofice2("{time}", "", wordDocument2);
                }
                if (checkBox3.Checked == false)
                {
                    ReplaceWord_Ofice2("{ИЗД№0000}", "ИЗД№" + nomer, wordDocument2);
                    ReplaceWord_Ofice2("{family}", Family, wordDocument2);
                    ReplaceWord_Ofice2("{date}", DATA_ss, wordDocument2);
                    ReplaceWord_Ofice2("{time}", Time_s, wordDocument2);
                }




                //wordDocument.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "PCB №" + nomer + "_" + DATA_ss + ".docx");
                wordDocument2.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "IZD №" + nomer + "_" + DATA_ss + ".docx");
                wordApp2.Visible = true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

       


        // Расчет 2-... для ... (для нового протокола (версия - 2) расчет ... для 2шт ...
        private void button21_Click(object sender, EventArgs e)  //... №1 Кнопка открыть РП1 (канал 1 ...) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox90.Text = ofd.FileName;
                dIRP1 = ofd.FileName;
                string prov_ofd = textBox90.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП1.txt" && dIRP1 != dIIRP1)
                {
                    textBox90.BackColor = System.Drawing.Color.LightGreen;
                    label213.Text = "";
                    IRP1 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП1.txt" ||  dIRP1 == dIIRP1)
                {
                    textBox90.BackColor = System.Drawing.Color.Red;
                    label213.Text = "некорректный файл";
                    IRP1 = 1;
                }
            }

            try
            {
                string path = textBox90.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA4Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB4Gh1 = Convert.ToDouble(res[3]);
                            textBox121.Text = IA4Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox126.Text = IB4Gh1.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA10Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB10Gh1 = Convert.ToDouble(res[3]);
                            textBox122.Text = IA10Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox127.Text = IB10Gh1.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA40Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB40Gh1 = Convert.ToDouble(res[3]);
                            textBox123.Text = IA40Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox128.Text = IB40Gh1.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA80Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB80Gh1 = Convert.ToDouble(res[3]);
                            textBox124.Text = IA80Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox129.Text = IB80Gh1.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA100Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB100Gh1 = Convert.ToDouble(res[3]);
                            textBox125.Text = IA100Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox130.Text = IB100Gh1.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button18_Click(object sender, EventArgs e)  // ... №1 Кнопка открыть РП2 (канал 2 ...) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox87.Text = ofd.FileName;
                dIRP2 = ofd.FileName;
                string prov_ofd = textBox87.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП2.txt" && dIRP2 != dIIRP2)
                {
                    textBox87.BackColor = System.Drawing.Color.LightGreen;
                    label212.Text = "";
                    IRP2 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП2.txt" || dIRP2 == dIIRP2)
                {
                    textBox87.BackColor = System.Drawing.Color.Red;
                    label212.Text = "некорректный файл";
                    IRP2 = 1;
                }
            }

            try
            {
                string path = textBox87.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA4Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB4Gh2 = Convert.ToDouble(res[3]);
                            textBox92.Text = IA4Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox97.Text = IB4Gh2.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA10Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB10Gh2 = Convert.ToDouble(res[3]);
                            textBox93.Text = IA10Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox98.Text = IB10Gh2.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA40Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB40Gh2 = Convert.ToDouble(res[3]);
                            textBox94.Text = IA40Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox99.Text = IB40Gh2.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA80Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB80Gh2 = Convert.ToDouble(res[3]);
                            textBox95.Text = IA80Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox100.Text = IB80Gh2.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA100Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB100Gh2 = Convert.ToDouble(res[3]);
                            textBox96.Text = IA100Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox101.Text = IB100Gh2.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button19_Click(object sender, EventArgs e)   // ...№1 Кнопка открыть РП3 (канал 3 ..) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox88.Text = ofd.FileName;
                dIRP3 = ofd.FileName;
                string prov_ofd = textBox88.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП3.txt" && dIRP3 != dIIRP3)
                {
                    textBox88.BackColor = System.Drawing.Color.LightGreen;
                    label214.Text = "";
                    IRP3 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП3.txt" || dIRP3 == dIIRP3)
                {
                    textBox88.BackColor = System.Drawing.Color.Red;
                    label214.Text = "некорректный файл";
                    IRP3 = 1;
                }
            }

            try
            {
                string path = textBox88.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA4Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB4Gh3 = Convert.ToDouble(res[3]);
                            textBox104.Text = IA4Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox109.Text = IB4Gh3.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA10Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB10Gh3 = Convert.ToDouble(res[3]);
                            textBox105.Text = IA10Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox110.Text = IB10Gh3.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA40Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB40Gh3 = Convert.ToDouble(res[3]);
                            textBox106.Text = IA40Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox111.Text = IB40Gh3.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA80Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB80Gh3 = Convert.ToDouble(res[3]);
                            textBox107.Text = IA80Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox112.Text = IB80Gh3.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA100Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB100Gh3 = Convert.ToDouble(res[3]);
                            textBox108.Text = IA100Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox113.Text = IB100Gh3.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button20_Click(object sender, EventArgs e)  // ... №1 Кнопка открыть РП4 (канал 4 ..) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox89.Text = ofd.FileName;
                dIRP4 = ofd.FileName;
                string prov_ofd = textBox89.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП4.txt" && dIRP4 != dIIRP4)
                {
                    textBox89.BackColor = System.Drawing.Color.LightGreen;
                    label215.Text = "";
                    IRP4 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП4.txt" || dIRP4 == dIIRP4)
                {
                    textBox89.BackColor = System.Drawing.Color.Red;
                    label215.Text = "некорректный файл";
                    IRP4 = 1;
                }
            }

            try
            {
                string path = textBox89.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA4Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB4Gh4 = Convert.ToDouble(res[3]);
                            textBox115.Text = IA4Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox120.Text = IB4Gh4.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA10Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB10Gh4 = Convert.ToDouble(res[3]);
                            textBox116.Text = IA10Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox132.Text = IB10Gh4.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA40Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB40Gh4 = Convert.ToDouble(res[3]);
                            textBox117.Text = IA40Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox133.Text = IB40Gh4.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA80Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB80Gh4 = Convert.ToDouble(res[3]);
                            textBox118.Text = IA80Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox134.Text = IB80Gh4.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IA100Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IB100Gh4 = Convert.ToDouble(res[3]);
                            textBox119.Text = IA100Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox135.Text = IB100Gh4.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        
        }

        private void button22_Click(object sender, EventArgs e)  // ...№2 Кнопка открыть РП1 (канал 1 ...П) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox102.Text = ofd.FileName;
                dIIRP1 = ofd.FileName;
                string prov_ofd = textBox102.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП1.txt" && dIRP1 != dIIRP1)
                {
                    textBox102.BackColor = System.Drawing.Color.LightGreen;
                    label216.Text = "";
                    IIRP1 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП1.txt" || dIRP1 == dIIRP1)
                {
                    textBox102.BackColor = System.Drawing.Color.Red;
                    label216.Text = "некорректный файл";
                    IIRP1 = 1;
                }
            }

            try
            {
                string path = textBox102.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA4Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB4Gh1 = Convert.ToDouble(res[3]);
                            textBox137.Text = IIA4Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox142.Text = IIB4Gh1.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA10Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB10Gh1 = Convert.ToDouble(res[3]);
                            textBox138.Text = IIA10Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox143.Text = IIB10Gh1.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA40Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB40Gh1 = Convert.ToDouble(res[3]);
                            textBox139.Text = IIA40Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox144.Text = IIB40Gh1.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA80Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB80Gh1 = Convert.ToDouble(res[3]);
                            textBox140.Text = IIA80Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox145.Text = IIB80Gh1.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA100Gh1 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB100Gh1 = Convert.ToDouble(res[3]);
                            textBox141.Text = IIA100Gh1.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox146.Text = IIB100Gh1.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button23_Click(object sender, EventArgs e)  // ... №2 Кнопка открыть РП2 (канал 2 ..) *txt
        {
             ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox103.Text = ofd.FileName;
                dIIRP2 = ofd.FileName;
                string prov_ofd = textBox103.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП2.txt" && dIRP2 != dIIRP2)
                {
                    textBox103.BackColor = System.Drawing.Color.LightGreen;
                    label217.Text = "";
                    IIRP2 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП2.txt" || dIRP2 == dIIRP2)
                {
                    textBox103.BackColor = System.Drawing.Color.Red;
                    label217.Text = "некорректный файл";
                    IIRP2 = 1;
                }
            }

            try
            {
                string path = textBox103.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA4Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB4Gh2 = Convert.ToDouble(res[3]);
                            textBox148.Text = IIA4Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox153.Text = IIB4Gh2.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA10Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB10Gh2 = Convert.ToDouble(res[3]);
                            textBox149.Text = IIA10Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox154.Text = IIB10Gh2.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA40Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB40Gh2 = Convert.ToDouble(res[3]);
                            textBox150.Text = IIA40Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox155.Text = IIB40Gh2.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA80Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB80Gh2 = Convert.ToDouble(res[3]);
                            textBox151.Text = IIA80Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox156.Text = IIB80Gh2.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA100Gh2 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB100Gh2 = Convert.ToDouble(res[3]);
                            textBox152.Text = IIA100Gh2.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox157.Text = IIB100Gh2.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        
        }

        private void button24_Click(object sender, EventArgs e)  // ..№2 Кнопка открыть РП3 (канал 3 ..) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox114.Text = ofd.FileName;
                dIIRP3 = ofd.FileName;
                string prov_ofd = textBox114.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП3.txt" && dIRP3 != dIIRP3)
                {
                    textBox114.BackColor = System.Drawing.Color.LightGreen;
                    label218.Text = "";
                    IIRP3 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП3.txt" || dIRP3 == dIIRP3)
                {
                    textBox114.BackColor = System.Drawing.Color.Red;
                    label218.Text = "некорректный файл";
                    IIRP3 = 1;
                }
            }

            try
            {
                string path = textBox114.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA4Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB4Gh3 = Convert.ToDouble(res[3]);
                            textBox159.Text = IIA4Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox164.Text = IIB4Gh3.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA10Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB10Gh3 = Convert.ToDouble(res[3]);
                            textBox160.Text = IIA10Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox165.Text = IIB10Gh3.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA40Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB40Gh3 = Convert.ToDouble(res[3]);
                            textBox161.Text = IIA40Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox166.Text = IIB40Gh3.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA80Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB80Gh3 = Convert.ToDouble(res[3]);
                            textBox162.Text = IIA80Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox167.Text = IIB80Gh3.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA100Gh3 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB100Gh3 = Convert.ToDouble(res[3]);
                            textBox163.Text = IIA100Gh3.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox168.Text = IIB100Gh3.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button25_Click(object sender, EventArgs e)  // .. №2 Кнопка открыть РП4 (канал 4 ..) *txt
        {
            ofd.Filter = "TXT|*.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox136.Text = ofd.FileName;
                dIIRP4 = ofd.FileName;
                string prov_ofd = textBox136.Text;
                string[] prov_ofd2 = prov_ofd.Split(new char[] { '\\' });
                if (prov_ofd2[prov_ofd2.Length - 1] == "Проверка АФЧХ ВКЗ_РП4.txt" && dIRP4 != dIIRP4)
                {
                    textBox136.BackColor = System.Drawing.Color.LightGreen;
                    label219.Text = "";
                    IIRP4 = 0;
                }
                if (prov_ofd2[prov_ofd2.Length - 1] != "Проверка АФЧХ ВКЗ_РП4.txt" || dIRP4 == dIIRP4)
                {
                    textBox136.BackColor = System.Drawing.Color.Red;
                    label219.Text = "некорректный файл";
                    IIRP4 = 1;
                }
            }

            try
            {
                string path = textBox136.Text;
                Console.WriteLine("Считываем посимвольно");
                using (StreamReader sr = new StreamReader(path, Encoding.Default))
                {
                    int total_txt = 0;
                    string line2;
                    string b;
                    while ((line2 = sr.ReadLine()) != null)
                    {
                        b = line2;
                        if (total_txt == 2) // поиск нужных нам цифр в строке 2 --- 4 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA4Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB4Gh4 = Convert.ToDouble(res[3]);
                            textBox170.Text = IIA4Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox175.Text = IIB4Gh4.ToString("f2");
                        }
                        if (total_txt == 3) // поиск нужных нам цифр в строке 3 --- 10 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA10Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB10Gh4 = Convert.ToDouble(res[3]);
                            textBox171.Text = IIA10Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox176.Text = IIB10Gh4.ToString("f2");
                        }
                        if (total_txt == 4) // поиск нужных нам цифр в строке 4 --- 40 ...
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA40Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB40Gh4 = Convert.ToDouble(res[3]);
                            textBox172.Text = IIA40Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox177.Text = IIB40Gh4.ToString("f2");
                        }
                        if (total_txt == 5) // поиск нужных нам цифр в строке 5 --- 80 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA80Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB80Gh4 = Convert.ToDouble(res[3]);
                            textBox173.Text = IIA80Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox178.Text = IIB80Gh4.ToString("f2");
                        }
                        if (total_txt == 6) // поиск нужных нам цифр в строке 6 --- 100 ..
                        {
                            string[] res = b.Split(new char[] { '\t' }); // разделяем полученную строку на массив символов разделенные табуляцией (или пробелами)
                            IIA100Gh4 = Convert.ToDouble(res[2]);   // преобразование типа строки в число с точкой
                            IIB100Gh4 = Convert.ToDouble(res[3]);
                            textBox174.Text = IIA100Gh4.ToString("f2"); // подставляем значение из txt в окно для альфа А
                            textBox179.Text = IIB100Gh4.ToString("f2");
                        }
                        ++total_txt;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void button27_Click(object sender, EventArgs e)  // Кнопка сброс
        {
            //.. №1
            // канал 1
            IA4Gh1 = -9999.9;
            IB4Gh1 = -9999.9;
            IA10Gh1 = -9999.9;
            IB10Gh1 = -9999.9;
            IA40Gh1 = -9999.9;
            IB40Gh1 = -9999.9;
            IA80Gh1 = -9999.9;
            IB80Gh1 = -9999.9;
            IA100Gh1 = -9999.9;
            IB100Gh1 = -9999.9;
            // канал 2
            IA4Gh2 = -9999.9;
            IB4Gh2 = -9999.9;
            IA10Gh2 = -9999.9;
            IB10Gh2 = -9999.9;
            IA40Gh2 = -9999.9;
            IB40Gh2 = -9999.9;
            IA80Gh2 = -9999.9;
            IB80Gh2 = -9999.9;
            IA100Gh2 = -9999.9;
            IB100Gh2 = -9999.9;
            // канал 3
            IA4Gh3 = -9999.9;
            IB4Gh3 = -9999.9;
            IA10Gh3 = -9999.9;
            IB10Gh3 = -9999.9;
            IA40Gh3 = -9999.9;
            IB40Gh3 = -9999.9;
            IA80Gh3 = -9999.9;
            IB80Gh3 = -9999.9;
            IA100Gh3 = -9999.9;
            IB100Gh3 = -9999.9;
            // канал 4
            IA4Gh4 = -9999.9;
            IB4Gh4 = -9999.9;
            IA10Gh4 = -9999.9;
            IB10Gh4 = -9999.9;
            IA40Gh4 = -9999.9;
            IB40Gh4 = -9999.9;
            IA80Gh4 = -9999.9;
            IB80Gh4 = -9999.9;
            IA100Gh4 = -9999.9;
            IB100Gh4 = -9999.9;

            //... №2
            // канал 1
            IIA4Gh1 = -9999.9;
            IIB4Gh1 = -9999.9;
            IIA10Gh1 = -9999.9;
            IIB10Gh1 = -9999.9;
            IIA40Gh1 = -9999.9;
            IIB40Gh1 = -9999.9;
            IIA80Gh1 = -9999.9;
            IIB80Gh1 = -9999.9;
            IIA100Gh1 = -9999.9;
            IIB100Gh1 = -9999.9;
            // канал 2
            IIA4Gh2 = -9999.9;
            IIB4Gh2 = -9999.9;
            IIA10Gh2 = -9999.9;
            IIB10Gh2 = -9999.9;
            IIA40Gh2 = -9999.9;
            IIB40Gh2 = -9999.9;
            IIA80Gh2 = -9999.9;
            IIB80Gh2 = -9999.9;
            IIA100Gh2 = -9999.9;
            IIB100Gh2 = -9999.9;
            // канал 3
            IIA4Gh3 = -9999.9;
            IIB4Gh3 = -9999.9;
            IIA10Gh3 = -9999.9;
            IIB10Gh3 = -9999.9;
            IIA40Gh3 = -9999.9;
            IIB40Gh3 = -9999.9;
            IIA80Gh3 = -9999.9;
            IIB80Gh3 = -9999.9;
            IIA100Gh3 = -9999.9;
            IIB100Gh3 = -9999.9;
            // канал 4
            IIA4Gh4 = -9999.9;
            IIB4Gh4 = -9999.9;
            IIA10Gh4 = -9999.9;
            IIB10Gh4 = -9999.9;
            IIA40Gh4 = -9999.9;
            IIB40Gh4 = -9999.9;
            IIA80Gh4 = -9999.9;
            IIB80Gh4 = -9999.9;
            IIA100Gh4 = -9999.9;
            IIB100Gh4 = -9999.9;

            dIRP1 = "1";  // адрес хранения файла "Проверка.. ВКЗ_РП1" для ..1
            dIRP2 = "2";
            dIRP3 = "3";
            dIRP4 = "4";
            dIIRP1 = "5"; // адрес хранения файла "Проверка ... ВКЗ_РП1" для ..2
            dIIRP2 = "6";
            dIIRP3 = "7";
            dIIRP4 = "8";

            // переменные проверки правильности загруженных txt файлов для канала1-4 (РП1-РП4.txt)
            IRP1 = 0;
            IIRP1 = 0;
            IRP2 = 0;
            IIRP2 = 0;
            IRP3 = 0;
            IIRP3 = 0;
            IRP4 = 0;
            IIRP4 = 0;

            // Требования ТУ
            label209.Text = "";
            label211.Text = "";
            label210.Text = "";
            label208.Text = "";
            label203.Text = "";
            label207.Text = "";
            label202.Text = "";
            label206.Text = "";
            label168.Text = "";
            label205.Text = "";
            label167.Text = "";
            label204.Text = "";
            label166.Text = "";
   
            //  .. №1 - 2 канал
            textBox90.Text = "";
            textBox90.BackColor = System.Drawing.Color.White; // выбор файла
            label213.Text = ""; // некорректный файл
            textBox121.Text = "";
            textBox121.BackColor = System.Drawing.Color.White;
            textBox122.Text = "";
            textBox122.BackColor = System.Drawing.Color.White;
            textBox123.Text = "";
            textBox123.BackColor = System.Drawing.Color.White;
            textBox124.Text = "";
            textBox124.BackColor = System.Drawing.Color.White;
            textBox125.Text = "";
            textBox125.BackColor = System.Drawing.Color.White;
            textBox126.Text = "";
            textBox126.BackColor = System.Drawing.Color.White;
            textBox127.Text = "";
            textBox127.BackColor = System.Drawing.Color.White;
            textBox128.Text = "";
            textBox128.BackColor = System.Drawing.Color.White;
            textBox129.Text = "";
            textBox129.BackColor = System.Drawing.Color.White;
            textBox130.Text = "";
            textBox130.BackColor = System.Drawing.Color.White;
            //  ... №1 - 2 канал
            textBox87.Text = "";
            textBox87.BackColor = System.Drawing.Color.White; // выбор файла
            label212.Text = ""; // некорректный файл
            textBox92.Text = "";
            textBox92.BackColor = System.Drawing.Color.White;
            textBox93.Text = "";
            textBox93.BackColor = System.Drawing.Color.White;
            textBox94.Text = "";
            textBox94.BackColor = System.Drawing.Color.White;
            textBox95.Text = "";
            textBox95.BackColor = System.Drawing.Color.White;
            textBox96.Text = "";
            textBox96.BackColor = System.Drawing.Color.White;
            textBox97.Text = "";
            textBox97.BackColor = System.Drawing.Color.White;
            textBox98.Text = "";
            textBox98.BackColor = System.Drawing.Color.White;
            textBox99.Text = "";
            textBox99.BackColor = System.Drawing.Color.White;
            textBox100.Text = "";
            textBox100.BackColor = System.Drawing.Color.White;
            textBox101.Text = "";
            textBox101.BackColor = System.Drawing.Color.White;
            //  ... №1 - 3 канал
            textBox88.Text = "";
            textBox88.BackColor = System.Drawing.Color.White; // выбор файла
            label214.Text = ""; // некорректный файл
            textBox104.Text = "";
            textBox104.BackColor = System.Drawing.Color.White;
            textBox105.Text = "";
            textBox105.BackColor = System.Drawing.Color.White;
            textBox106.Text = "";
            textBox106.BackColor = System.Drawing.Color.White;
            textBox107.Text = "";
            textBox107.BackColor = System.Drawing.Color.White;
            textBox108.Text = "";
            textBox108.BackColor = System.Drawing.Color.White;
            textBox109.Text = "";
            textBox109.BackColor = System.Drawing.Color.White;
            textBox110.Text = "";
            textBox110.BackColor = System.Drawing.Color.White;
            textBox111.Text = "";
            textBox111.BackColor = System.Drawing.Color.White;
            textBox112.Text = "";
            textBox112.BackColor = System.Drawing.Color.White;
            textBox113.Text = "";
            textBox113.BackColor = System.Drawing.Color.White;
            //  ... №1 - 4 канал
            textBox89.Text = "";
            textBox89.BackColor = System.Drawing.Color.White; // выбор файла
            label215.Text = ""; // некорректный файл
            textBox115.Text = "";
            textBox115.BackColor = System.Drawing.Color.White;
            textBox116.Text = "";
            textBox116.BackColor = System.Drawing.Color.White;
            textBox117.Text = "";
            textBox117.BackColor = System.Drawing.Color.White;
            textBox118.Text = "";
            textBox118.BackColor = System.Drawing.Color.White;
            textBox119.Text = "";
            textBox119.BackColor = System.Drawing.Color.White;
            textBox120.Text = "";
            textBox120.BackColor = System.Drawing.Color.White;
            textBox132.Text = "";
            textBox132.BackColor = System.Drawing.Color.White;
            textBox133.Text = "";
            textBox133.BackColor = System.Drawing.Color.White;
            textBox134.Text = "";
            textBox134.BackColor = System.Drawing.Color.White;
            textBox135.Text = "";
            textBox135.BackColor = System.Drawing.Color.White;

            //  ... №2 - 1 канал
            textBox102.Text = "";
            textBox102.BackColor = System.Drawing.Color.White; // выбор файла
            label216.Text = ""; // некорректный файл
            textBox137.Text = "";
            textBox137.BackColor = System.Drawing.Color.White;
            textBox138.Text = "";
            textBox138.BackColor = System.Drawing.Color.White;
            textBox139.Text = "";
            textBox139.BackColor = System.Drawing.Color.White;
            textBox140.Text = "";
            textBox140.BackColor = System.Drawing.Color.White;
            textBox141.Text = "";
            textBox141.BackColor = System.Drawing.Color.White;
            textBox142.Text = "";
            textBox142.BackColor = System.Drawing.Color.White;
            textBox143.Text = "";
            textBox143.BackColor = System.Drawing.Color.White;
            textBox144.Text = "";
            textBox144.BackColor = System.Drawing.Color.White;
            textBox145.Text = "";
            textBox145.BackColor = System.Drawing.Color.White;
            textBox146.Text = "";
            textBox146.BackColor = System.Drawing.Color.White;
            //  ... №2 - 2 канал
            textBox103.Text = "";
            textBox103.BackColor = System.Drawing.Color.White; // выбор файла
            label217.Text = ""; // некорректный файл
            textBox148.Text = "";
            textBox148.BackColor = System.Drawing.Color.White;
            textBox149.Text = "";
            textBox149.BackColor = System.Drawing.Color.White;
            textBox150.Text = "";
            textBox150.BackColor = System.Drawing.Color.White;
            textBox151.Text = "";
            textBox151.BackColor = System.Drawing.Color.White;
            textBox152.Text = "";
            textBox152.BackColor = System.Drawing.Color.White;
            textBox153.Text = "";
            textBox153.BackColor = System.Drawing.Color.White;
            textBox154.Text = "";
            textBox154.BackColor = System.Drawing.Color.White;
            textBox155.Text = "";
            textBox155.BackColor = System.Drawing.Color.White;
            textBox156.Text = "";
            textBox156.BackColor = System.Drawing.Color.White;
            textBox157.Text = "";
            textBox157.BackColor = System.Drawing.Color.White;
            //  ... №2 - 3 канал
            textBox114.Text = "";
            textBox114.BackColor = System.Drawing.Color.White; // выбор файла
            label218.Text = ""; // некорректный файл
            textBox159.Text = "";
            textBox159.BackColor = System.Drawing.Color.White;
            textBox160.Text = "";
            textBox160.BackColor = System.Drawing.Color.White;
            textBox161.Text = "";
            textBox161.BackColor = System.Drawing.Color.White;
            textBox162.Text = "";
            textBox162.BackColor = System.Drawing.Color.White;
            textBox163.Text = "";
            textBox163.BackColor = System.Drawing.Color.White;
            textBox164.Text = "";
            textBox164.BackColor = System.Drawing.Color.White;
            textBox165.Text = "";
            textBox165.BackColor = System.Drawing.Color.White;
            textBox166.Text = "";
            textBox166.BackColor = System.Drawing.Color.White;
            textBox167.Text = "";
            textBox167.BackColor = System.Drawing.Color.White;
            textBox168.Text = "";
            textBox168.BackColor = System.Drawing.Color.White;
            // ... №2 - 4 канал
            textBox136.Text = "";
            textBox136.BackColor = System.Drawing.Color.White; // выбор файла
            label219.Text = ""; // некорректный файл
            textBox170.Text = "";
            textBox170.BackColor = System.Drawing.Color.White;
            textBox171.Text = "";
            textBox171.BackColor = System.Drawing.Color.White;
            textBox172.Text = "";
            textBox172.BackColor = System.Drawing.Color.White;
            textBox173.Text = "";
            textBox173.BackColor = System.Drawing.Color.White;
            textBox174.Text = "";
            textBox174.BackColor = System.Drawing.Color.White;
            textBox175.Text = "";
            textBox175.BackColor = System.Drawing.Color.White;
            textBox176.Text = "";
            textBox176.BackColor = System.Drawing.Color.White;
            textBox177.Text = "";
            textBox177.BackColor = System.Drawing.Color.White;
            textBox178.Text = "";
            textBox178.BackColor = System.Drawing.Color.White;
            textBox179.Text = "";
            textBox179.BackColor = System.Drawing.Color.White;
            
        }

        private void button28_Click(object sender, EventArgs e)  // Кнопка проверка (проверяем значения ... для 2ух ..., что бы были в диапозоне требований ТУ
        {
            label209.Text = "Требования ТУ";
            label211.Text = "#";
            label210.Text = "#";
            label208.Text = A4Gh_Down + "..." + A4Gh_Up; ;
            label203.Text = B4Gh_Down + "..." + B4Gh_Up;
            label207.Text = A10Gh_Down + "..." + A10Gh_Up;
            label202.Text = B10Gh_Down + "..." + B10Gh_Up;
            label206.Text = A40Gh_Down + "..." + A40Gh_Up;
            label168.Text = B40Gh_Down + "..." + B40Gh_Up;
            label205.Text = A80Gh_Down + "..." + A80Gh_Up;
            label167.Text = B80Gh_Down + "..." + B80Gh_Up;
            label204.Text = A100Gh_Down + "..." + A100Gh_Up;
            label166.Text = B100Gh_Down + "..." + B100Gh_Up;

            // .. №1
            // Окрашиваем текст бокс в зеленый или красный в зависимости от вхождения в диапозон
            // .. - 4 .. канал 1-4
            if (A4Gh_Down <= IA4Gh1 && IA4Gh1 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox121.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox121.BackColor = System.Drawing.Color.Red;
            }
            if (A4Gh_Down <= IA4Gh2 && IA4Gh2 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox92.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox92.BackColor = System.Drawing.Color.Red;
            }
            if (A4Gh_Down <= IA4Gh3 && IA4Gh3 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox104.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox104.BackColor = System.Drawing.Color.Red;
            }
            if (A4Gh_Down <= IA4Gh4 && IA4Gh4 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox115.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox115.BackColor = System.Drawing.Color.Red;
            }

            // ..- 10 .. канал 1-4
            if (A10Gh_Down <= IA10Gh1 && IA10Gh1 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox122.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox122.BackColor = System.Drawing.Color.Red;
            }
            if (A10Gh_Down <= IA10Gh2 && IA10Gh2 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox93.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox93.BackColor = System.Drawing.Color.Red;
            }
            if (A10Gh_Down <= IA10Gh3 && IA10Gh3 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox105.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox105.BackColor = System.Drawing.Color.Red;
            }
            if (A10Gh_Down <= IA10Gh4 && IA10Gh4 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox116.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox116.BackColor = System.Drawing.Color.Red;
            }

            // .. - 40 ... канал 1-4
            if (A40Gh_Down <= IA40Gh1 && IA40Gh1 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox123.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox123.BackColor = System.Drawing.Color.Red;
            }
            if (A40Gh_Down <= IA40Gh2 && IA40Gh2 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox94.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox94.BackColor = System.Drawing.Color.Red;
            }
            if (A40Gh_Down <= IA40Gh3 && IA40Gh3 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox106.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox106.BackColor = System.Drawing.Color.Red;
            }
            if (A40Gh_Down <= IA40Gh4 && IA40Gh4 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox117.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox117.BackColor = System.Drawing.Color.Red;
            }

            // .. - 80 ... канал 1-4
            if (A80Gh_Down <= IA80Gh1 && IA80Gh1 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox124.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox124.BackColor = System.Drawing.Color.Red;
            }
            if (A80Gh_Down <= IA80Gh2 && IA80Gh2 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox95.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox95.BackColor = System.Drawing.Color.Red;
            }
            if (A80Gh_Down <= IA80Gh3 && IA80Gh3 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox107.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox107.BackColor = System.Drawing.Color.Red;
            }
            if (A80Gh_Down <= IA80Gh4 && IA80Gh4 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox118.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox118.BackColor = System.Drawing.Color.Red;
            }

            // .. - 100.. канал 1-4
            if (A100Gh_Down <= IA100Gh1 && IA100Gh1 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox125.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox125.BackColor = System.Drawing.Color.Red;
            }
            if (A100Gh_Down <= IA100Gh2 && IA100Gh2 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox96.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox96.BackColor = System.Drawing.Color.Red;
            }
            if (A100Gh_Down <= IA100Gh3 && IA100Gh3 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox108.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox108.BackColor = System.Drawing.Color.Red;
            }
            if (A100Gh_Down <= IA100Gh4 && IA100Gh4 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox119.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox119.BackColor = System.Drawing.Color.Red;
            }


            // .. - 4 .. канал 1-4
            if (B4Gh_Down <= IB4Gh1 && IB4Gh1 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox126.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox126.BackColor = System.Drawing.Color.Red;
            }
            if (B4Gh_Down <= IB4Gh2 && IB4Gh2 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox97.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox97.BackColor = System.Drawing.Color.Red;
            }
            if (B4Gh_Down <= IB4Gh3 && IB4Gh3 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox109.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox109.BackColor = System.Drawing.Color.Red;
            }
            if (B4Gh_Down <= IB4Gh4 && IB4Gh4 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox120.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox120.BackColor = System.Drawing.Color.Red;
            }

            // .. - 10 ... канал 1-4
            if (B10Gh_Down <= IB10Gh1 && IB10Gh1 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox127.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox127.BackColor = System.Drawing.Color.Red;
            }
            if (B10Gh_Down <= IB10Gh2 && IB10Gh2 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox98.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox98.BackColor = System.Drawing.Color.Red;
            }
            if (B10Gh_Down <= IB10Gh3 && IB10Gh3 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox110.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox110.BackColor = System.Drawing.Color.Red;
            }
            if (B10Gh_Down <= IB10Gh4 && IB10Gh4 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox132.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox132.BackColor = System.Drawing.Color.Red;
            }

            //.. - 40 .. канал 1-4
            if (B40Gh_Down <= IB40Gh1 && IB40Gh1 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox128.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox128.BackColor = System.Drawing.Color.Red;
            }
            if (B40Gh_Down <= IB40Gh2 && IB40Gh2 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox99.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox99.BackColor = System.Drawing.Color.Red;
            }
            if (B40Gh_Down <= IB40Gh3 && IB40Gh3 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox111.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox111.BackColor = System.Drawing.Color.Red;
            }
            if (B40Gh_Down <= IB40Gh4 && IB40Gh4 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox133.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox133.BackColor = System.Drawing.Color.Red;
            }

            // .. - 80 .. канал 1-4
            if (B80Gh_Down <= IB80Gh1 && IB80Gh1 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox129.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox129.BackColor = System.Drawing.Color.Red;
            }
            if (B80Gh_Down <= IB80Gh2 && IB80Gh2 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox100.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox100.BackColor = System.Drawing.Color.Red;
            }
            if (B80Gh_Down <= IB80Gh3 && IB80Gh3 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox112.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox112.BackColor = System.Drawing.Color.Red;
            }
            if (B80Gh_Down <= IB80Gh4 && IB80Gh4 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox134.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox134.BackColor = System.Drawing.Color.Red;
            }

            //.. - 100 ... канал 1-4
            if (B100Gh_Down <= IB100Gh1 && IB100Gh1 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox130.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox130.BackColor = System.Drawing.Color.Red;
            }
            if (B100Gh_Down <= IB100Gh2 && IB100Gh2 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox101.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox101.BackColor = System.Drawing.Color.Red;
            }
            if (B100Gh_Down <= IB100Gh3 && IB100Gh3 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox113.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox113.BackColor = System.Drawing.Color.Red;
            }
            if (B100Gh_Down <= IB100Gh4 && IB100Gh4 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
            {
                textBox135.BackColor = System.Drawing.Color.LightGreen;
            }
            else
            {
                textBox135.BackColor = System.Drawing.Color.Red;
            }

            if (checkBox5.Checked == false)
            {
                // .. №2
                // Окрашиваем текст бокс в зеленый или красный в зависимости от вхождения в диапозон
                // .. - 4 .. канал 1-4
                if (A4Gh_Down <= IIA4Gh1 && IIA4Gh1 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox137.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox137.BackColor = System.Drawing.Color.Red;
                }
                if (A4Gh_Down <= IIA4Gh2 && IIA4Gh2 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox148.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox148.BackColor = System.Drawing.Color.Red;
                }
                if (A4Gh_Down <= IIA4Gh3 && IIA4Gh3 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox159.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox159.BackColor = System.Drawing.Color.Red;
                }
                if (A4Gh_Down <= IIA4Gh4 && IIA4Gh4 <= A4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox170.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox170.BackColor = System.Drawing.Color.Red;
                }

                // .. - 10.. канал 1-4
                if (A10Gh_Down <= IIA10Gh1 && IIA10Gh1 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox138.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox138.BackColor = System.Drawing.Color.Red;
                }
                if (A10Gh_Down <= IIA10Gh2 && IIA10Gh2 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox149.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox149.BackColor = System.Drawing.Color.Red;
                }
                if (A10Gh_Down <= IIA10Gh3 && IIA10Gh3 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox160.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox160.BackColor = System.Drawing.Color.Red;
                }
                if (A10Gh_Down <= IIA10Gh4 && IIA10Gh4 <= A10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox171.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox171.BackColor = System.Drawing.Color.Red;
                }

                // .. - 40 .. канал 1-4
                if (A40Gh_Down <= IIA40Gh1 && IIA40Gh1 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox139.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox139.BackColor = System.Drawing.Color.Red;
                }
                if (A40Gh_Down <= IIA40Gh2 && IIA40Gh2 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox150.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox150.BackColor = System.Drawing.Color.Red;
                }
                if (A40Gh_Down <= IIA40Gh3 && IIA40Gh3 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox161.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox161.BackColor = System.Drawing.Color.Red;
                }
                if (A40Gh_Down <= IIA40Gh4 && IIA40Gh4 <= A40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox172.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox172.BackColor = System.Drawing.Color.Red;
                }

                // .. - 80 .. канал 1-4
                if (A80Gh_Down <= IIA80Gh1 && IIA80Gh1 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox140.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox140.BackColor = System.Drawing.Color.Red;
                }
                if (A80Gh_Down <= IIA80Gh2 && IIA80Gh2 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox151.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox151.BackColor = System.Drawing.Color.Red;
                }
                if (A80Gh_Down <= IIA80Gh3 && IIA80Gh3 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox162.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox162.BackColor = System.Drawing.Color.Red;
                }
                if (A80Gh_Down <= IIA80Gh4 && IIA80Gh4 <= A80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox173.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox173.BackColor = System.Drawing.Color.Red;
                }

                // .. - 100 ... канал 1-4
                if (A100Gh_Down <= IIA100Gh1 && IIA100Gh1 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox141.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox141.BackColor = System.Drawing.Color.Red;
                }
                if (A100Gh_Down <= IIA100Gh2 && IIA100Gh2 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox152.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox152.BackColor = System.Drawing.Color.Red;
                }
                if (A100Gh_Down <= IIA100Gh3 && IIA100Gh3 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox163.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox163.BackColor = System.Drawing.Color.Red;
                }
                if (A100Gh_Down <= IIA100Gh4 && IIA100Gh4 <= A100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox174.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox174.BackColor = System.Drawing.Color.Red;
                }


                // .. - 4 .. канал 1-4
                if (B4Gh_Down <= IIB4Gh1 && IIB4Gh1 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox142.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox142.BackColor = System.Drawing.Color.Red;
                }
                if (B4Gh_Down <= IIB4Gh2 && IIB4Gh2 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox153.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox153.BackColor = System.Drawing.Color.Red;
                }
                if (B4Gh_Down <= IIB4Gh3 && IIB4Gh3 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox164.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox164.BackColor = System.Drawing.Color.Red;
                }
                if (B4Gh_Down <= IIB4Gh4 && IIB4Gh4 <= B4Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox175.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox175.BackColor = System.Drawing.Color.Red;
                }

                // .. - 10 .. канал 1-4
                if (B10Gh_Down <= IIB10Gh1 && IIB10Gh1 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox143.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox143.BackColor = System.Drawing.Color.Red;
                }
                if (B10Gh_Down <= IIB10Gh2 && IIB10Gh2 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox154.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox154.BackColor = System.Drawing.Color.Red;
                }
                if (B10Gh_Down <= IIB10Gh3 && IIB10Gh3 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox165.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox165.BackColor = System.Drawing.Color.Red;
                }
                if (B10Gh_Down <= IIB10Gh4 && IIB10Gh4 <= B10Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox176.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox176.BackColor = System.Drawing.Color.Red;
                }

                // .. - 40 .. канал 1-4
                if (B40Gh_Down <= IIB40Gh1 && IIB40Gh1 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox144.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox144.BackColor = System.Drawing.Color.Red;
                }
                if (B40Gh_Down <= IIB40Gh2 && IIB40Gh2 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox155.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox155.BackColor = System.Drawing.Color.Red;
                }
                if (B40Gh_Down <= IIB40Gh3 && IIB40Gh3 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox166.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox166.BackColor = System.Drawing.Color.Red;
                }
                if (B40Gh_Down <= IIB40Gh4 && IIB40Gh4 <= B40Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox177.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox177.BackColor = System.Drawing.Color.Red;
                }

                // .. - 80... канал 1-4
                if (B80Gh_Down <= IIB80Gh1 && IIB80Gh1 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox145.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox145.BackColor = System.Drawing.Color.Red;
                }
                if (B80Gh_Down <= IIB80Gh2 && IIB80Gh2 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox156.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox156.BackColor = System.Drawing.Color.Red;
                }
                if (B80Gh_Down <= IIB80Gh3 && IIB80Gh3 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox167.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox167.BackColor = System.Drawing.Color.Red;
                }
                if (B80Gh_Down <= IIB80Gh4 && IIB80Gh4 <= B80Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox178.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox178.BackColor = System.Drawing.Color.Red;
                }

                // .. - 100 ... канал 1-4
                if (B100Gh_Down <= IIB100Gh1 && IIB100Gh1 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox146.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox146.BackColor = System.Drawing.Color.Red;
                }
                if (B100Gh_Down <= IIB100Gh2 && IIB100Gh2 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox157.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox157.BackColor = System.Drawing.Color.Red;
                }
                if (B100Gh_Down <= IIB100Gh3 && IIB100Gh3 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox168.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox168.BackColor = System.Drawing.Color.Red;
                }
                if (B100Gh_Down <= IIB100Gh4 && IIB100Gh4 <= B100Gh_Up) // проверяем полученное значение с требуемым диапозоном если входит то текс бокс окрашиваем в зеленый, если нет то в красный
                {
                    textBox179.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    textBox179.BackColor = System.Drawing.Color.Red;
                }
            }
        }

        private void button26_Click(object sender, EventArgs e)   // Вывести протокол проверки 2-.. в Microsoft Word (для 2шт ..)
        {
            var wordApp3 = new Word.Application();//создаем переменную "wordApp3" с приложением оболочки ворда
            wordApp3.Visible = false; //не видеть в процессе экспорта открытое окно ворда

            try
            {
                var nomer = textBox86.Text; // номер ..
                var Family = textBox22.Text; // Фамилия И.О. проверяющего  
                string DATA_s = System.DateTime.Now.ToShortDateString(); //дата сокращенная
                string Time_s = System.DateTime.Now.ToShortTimeString(); //время короткое
                string DATA_ss = System.DateTime.Now.Year.ToString() + "." + System.DateTime.Now.Month.ToString("d2") + "." + System.DateTime.Now.Day.ToString();
                // работа с шаблоном
                var wordDocument3 = wordApp3.Documents.Open(TemplateFileName3);//открываем документ


                // меняем форматирование текста в заисимости от полученных результатов (подсвечиваем красным если вне диапозона) ред. 2021.04.07
                // взято с ресурса https://fooobar.com/questions/516118/c-searching-a-text-in-word-and-getting-the-range-of-the-result
                Word.Range range;
                Word.Range temprange;
                Word.Selection currentSelection;
                
                // .. №1
                // ..4 ..
                if (A4Gh_Down > IA4Gh1 || IA4Gh1 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA4Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA4Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > IA4Gh2 || IA4Gh2 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA4Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA4Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > IA4Gh3 || IA4Gh3 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA4Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA4Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > IA4Gh4 || IA4Gh4 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA4Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA4Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 10 ..
                if (A10Gh_Down > IA10Gh1 || IA10Gh1 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA10Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA10Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > IA10Gh2 || IA10Gh2 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA10Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA10Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > IA10Gh3 || IA10Gh3 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA10Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA10Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > IA10Gh4 || IA10Gh4 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA10Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA10Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 40 ..
                if (A40Gh_Down > IA40Gh1 || IA40Gh1 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA40Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA40Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > IA40Gh2 || IA40Gh2 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA40Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA40Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > IA40Gh3 || IA40Gh3 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA40Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA40Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > IA40Gh4 || IA40Gh4 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA40Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA40Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 80 ..
                if (A80Gh_Down > IA80Gh1 || IA80Gh1 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA80Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA80Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > IA80Gh2 || IA80Gh2 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA80Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA80Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > IA80Gh3 || IA80Gh3 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA80Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA80Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > IA80Gh4 || IA80Gh4 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA80Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA80Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 100...
                if (A100Gh_Down > IA100Gh1 || IA100Gh1 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA100Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA100Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > IA100Gh2 || IA100Gh2 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA100Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA100Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > IA100Gh3 || IA100Gh3 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA100Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA100Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > IA100Gh4 || IA100Gh4 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IA100Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IA100Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }


                // .. 4 ...
                if (B4Gh_Down > IB4Gh1 || IB4Gh1 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB4Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB4Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > IB4Gh2 || IB4Gh2 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB4Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB4Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > IB4Gh3 || IB4Gh3 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB4Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB4Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > IB4Gh4 || IB4Gh4 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB4Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB4Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 10 ...
                if (B10Gh_Down > IB10Gh1 || IB10Gh1 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB10Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB10Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > IB10Gh2 || IB10Gh2 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB10Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB10Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > IB10Gh3 || IB10Gh3 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB10Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB10Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > IB10Gh4 || IB10Gh4 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB10Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB10Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 40 ...
                if (B40Gh_Down > IB40Gh1 || IB40Gh1 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB40Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB40Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > IB40Gh2 || IB40Gh2 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB40Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB40Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > IB40Gh3 || IB40Gh3 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB40Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB40Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > IB40Gh4 || IB40Gh4 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB40Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB40Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 80 ..
                if (B80Gh_Down > IB80Gh1 || IB80Gh1 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB80Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB80Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > IB80Gh2 || IB80Gh2 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB80Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB80Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > IB80Gh3 || IB80Gh3 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB80Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB80Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > IB80Gh4 || IB80Gh4 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB80Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB80Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                //.. 100 ..
                if (B100Gh_Down > IB100Gh1 || IB100Gh1 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB100Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB100Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > IB100Gh2 || IB100Gh2 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB100Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB100Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > IB100Gh3 || IB100Gh3 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB100Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB100Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > IB100Gh4 || IB100Gh4 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IB100Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IB100Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }


                // .. №2
                // .. 4 ...
                if (A4Gh_Down > IIA4Gh1 || IIA4Gh1 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA4Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA4Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > IIA4Gh2 || IIA4Gh2 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA4Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA4Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > IIA4Gh3 || IIA4Gh3 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA4Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA4Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A4Gh_Down > IIA4Gh4 || IIA4Gh4 > A4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA4Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA4Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 10 ...
                if (A10Gh_Down > IIA10Gh1 || IIA10Gh1 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA10Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA10Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > IIA10Gh2 || IIA10Gh2 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA10Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA10Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > IIA10Gh3 || IIA10Gh3 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA10Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA10Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A10Gh_Down > IIA10Gh4 || IIA10Gh4 > A10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA10Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA10Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 40 ..
                if (A40Gh_Down > IIA40Gh1 || IIA40Gh1 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA40Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA40Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > IIA40Gh2 || IIA40Gh2 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA40Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA40Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > IIA40Gh3 || IIA40Gh3 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA40Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA40Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A40Gh_Down > IIA40Gh4 || IIA40Gh4 > A40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA40Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA40Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 80 ..
                if (A80Gh_Down > IIA80Gh1 || IIA80Gh1 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA80Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA80Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > IIA80Gh2 || IIA80Gh2 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA80Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA80Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > IIA80Gh3 || IIA80Gh3 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA80Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA80Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A80Gh_Down > IIA80Gh4 || IIA80Gh4 > A80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA80Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA80Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                //.. 100 ..
                if (A100Gh_Down > IIA100Gh1 || IIA100Gh1 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA100Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA100Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > IIA100Gh2 || IIA100Gh2 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA100Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA100Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > IIA100Gh3 || IIA100Gh3 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA100Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA100Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (A100Gh_Down > IIA100Gh4 || IIA100Gh4 > A100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIA100Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIA100Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }


                // .. 4 ..
                if (B4Gh_Down > IIB4Gh1 || IIB4Gh1 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB4Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB4Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > IIB4Gh2 || IIB4Gh2 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB4Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB4Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > IIB4Gh3 || IIB4Gh3 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB4Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB4Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B4Gh_Down > IIB4Gh4 || IIB4Gh4 > B4Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB4Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB4Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 10 ...
                if (B10Gh_Down > IIB10Gh1 || IIB10Gh1 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB10Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB10Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > IIB10Gh2 || IIB10Gh2 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB10Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB10Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > IIB10Gh3 || IIB10Gh3 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB10Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB10Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B10Gh_Down > IIB10Gh4 || IIB10Gh4 > B10Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB10Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB10Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 40 ...
                if (B40Gh_Down > IIB40Gh1 || IIB40Gh1 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB40Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB40Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > IIB40Gh2 || IIB40Gh2 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB40Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB40Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > IIB40Gh3 || IIB40Gh3 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB40Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB40Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B40Gh_Down > IIB40Gh4 || IIB40Gh4 > B40Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB40Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB40Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 80 ...
                if (B80Gh_Down > IIB80Gh1 || IIB80Gh1 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB80Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB80Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > IIB80Gh2 || IIB80Gh2 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB80Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB80Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > IIB80Gh3 || IIB80Gh3 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB80Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB80Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B80Gh_Down > IIB80Gh4 || IIB80Gh4 > B80Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB80Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB80Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // .. 100 ..
                if (B100Gh_Down > IIB100Gh1 || IIB100Gh1 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB100Gh1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB100Gh1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > IIB100Gh2 || IIB100Gh2 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB100Gh2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB100Gh2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > IIB100Gh3 || IIB100Gh3 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB100Gh3}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB100Gh3}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (B100Gh_Down > IIB100Gh4 || IIB100Gh4 > B100Gh_Up)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{IIB100Gh4}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{IIB100Gh4}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }


                // проверки правильности загруженных txt файлов для канала1-4 (РП1-РП4.txt)
                // Если расчитан ... для канала1 не из файла "Проверка .. ВКЗ_РП1"  то в протоколе канал 1 подсвечиваеся красным и тд для остальных каналов
                // Для ..№1
                if (IRP1 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 1");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 1"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (IRP2 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 2");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 2"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (IRP3 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 3");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 3"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (IRP4 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 4");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 4"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД1}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД1}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }

                // Для ...2
                if (IIRP1 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 1");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 1"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (IIRP2 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 2");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 2"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (IIRP3 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 3");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 3"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }
                if (IIRP4 == 1)
                {
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("канал 4");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("канал 4"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                    wordApp3.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue; //
                    wordApp3.Selection.Find.Execute("{№ИЗД2}");
                    range = wordApp3.Selection.Range;
                    if (range.Text.Contains("{№ИЗД2}"))
                    {
                        // gets desired range here it gets last character to make superscript in range
                        temprange = wordDocument3.Range(range.Start, range.End);
                        temprange.Select();
                        currentSelection = wordApp3.Selection;
                        currentSelection.FormattedText.HighlightColorIndex = Word.WdColorIndex.wdRed;
                    }
                }


                if (IB4Gh1 == -9999.9 && IB100Gh3 == -9999.9 && IB80Gh4 == -9999.9) //если файл не выбран
                {
                    ReplaceWord_Ofice2("{IA4Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA4Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA4Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA4Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh4}", "", wordDocument3);
                    // Для ...№2 поиск и замена для .. 
                    ReplaceWord_Ofice2("{IB4Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB4Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB4Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB4Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh4}", "", wordDocument3);
                }

                else //если файл выбран и подставлены значения
                {
                    // Для ...№1 поиск и замена для ...
                    ReplaceWord_Ofice2("{IA4Gh1}", IA4Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA4Gh2}", IA4Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA4Gh3}", IA4Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA4Gh4}", IA4Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh1}", IA10Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh2}", IA10Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh3}", IA10Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA10Gh4}", IA10Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh1}", IA40Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh2}", IA40Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh3}", IA40Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA40Gh4}", IA40Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh1}", IA80Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh2}", IA80Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh3}", IA80Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA80Gh4}", IA80Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh1}", IA100Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh2}", IA100Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh3}", IA100Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IA100Gh4}", IA100Gh4.ToString(), wordDocument3);
                    // Для ..№1 поиск и замена для ...
                    ReplaceWord_Ofice2("{IB4Gh1}", IB4Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB4Gh2}", IB4Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB4Gh3}", IB4Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB4Gh4}", IB4Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh1}", IB10Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh2}", IB10Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh3}", IB10Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB10Gh4}", IB10Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh1}", IB40Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh2}", IB40Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh3}", IB40Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB40Gh4}", IB40Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh1}", IB80Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh2}", IB80Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh3}", IB80Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB80Gh4}", IB80Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh1}", IB100Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh2}", IB100Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh3}", IB100Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IB100Gh4}", IB100Gh4.ToString(), wordDocument3);
                }

                if (checkBox5.Checked == false && IIB4Gh1 != -9999.9 && IIB100Gh3 != -9999.9) // Для ..№2 ...
                {
                    // Для ..№2 поиск и замена для ..
                    ReplaceWord_Ofice2("{IIA4Gh1}", IIA4Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA4Gh2}", IIA4Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA4Gh3}", IIA4Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA4Gh4}", IIA4Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh1}", IIA10Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh2}", IIA10Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh3}", IIA10Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh4}", IIA10Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh1}", IIA40Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh2}", IIA40Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh3}", IIA40Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh4}", IIA40Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh1}", IIA80Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh2}", IIA80Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh3}", IIA80Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh4}", IIA80Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh1}", IIA100Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh2}", IIA100Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh3}", IIA100Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh4}", IIA100Gh4.ToString(), wordDocument3);
                    // Для ...№2 поиск и замена для ...
                    ReplaceWord_Ofice2("{IIB4Gh1}", IIB4Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB4Gh2}", IIB4Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB4Gh3}", IIB4Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB4Gh4}", IIB4Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh1}", IIB10Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh2}", IIB10Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh3}", IIB10Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh4}", IIB10Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh1}", IIB40Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh2}", IIB40Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh3}", IIB40Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh4}", IIB40Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh1}", IIB80Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh2}", IIB80Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh3}", IIB80Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh4}", IIB80Gh4.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh1}", IIB100Gh1.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh2}", IIB100Gh2.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh3}", IIB100Gh3.ToString(), wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh4}", IIB100Gh4.ToString(), wordDocument3);
                }

                if (checkBox5.Checked == true || (IIB4Gh1 == -9999.9 && IIB100Gh3 == -9999.9))
                {
                    ReplaceWord_Ofice2("{IIA4Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA4Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA4Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA4Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA10Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA40Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA80Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIA100Gh4}", "", wordDocument3);
                    // Для ..№2 поиск и замена для ...
                    ReplaceWord_Ofice2("{IIB4Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB4Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB4Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB4Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB10Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB40Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB80Gh4}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh3}", "", wordDocument3);
                    ReplaceWord_Ofice2("{IIB100Gh4}", "", wordDocument3);
                }

                if (checkBox4.Checked == true) // Отменить вставку ...№/Фамилия/Дата/Время
                {
                    ReplaceWord_Ofice2("{№ИЗД1}", "", wordDocument3);
                    ReplaceWord_Ofice2("{№ИЗД2}", "", wordDocument3);
                    ReplaceWord_Ofice2("{family}", "", wordDocument3);
                    ReplaceWord_Ofice2("{date}", "", wordDocument3);
                    ReplaceWord_Ofice2("{time}", "", wordDocument3);
                }

                if (checkBox4.Checked == false && checkBox6.Checked == false) // По умолчанию вставляем в протокол ...№/Фамилия/Дата/Время
                {
                    if (checkBox5.Checked == true)
                    {
                        ReplaceWord_Ofice2("{№ИЗД2}", "", wordDocument3);
                    }
                    if (checkBox5.Checked == false)
                    {
                        ReplaceWord_Ofice2("{№ИЗД2}", "ИЗД №" + textBox91.Text, wordDocument3);
                    }
                    ReplaceWord_Ofice2("{№ИЗД1}", "ИЗД №" + textBox131.Text, wordDocument3);
                    ReplaceWord_Ofice2("{family}", Family, wordDocument3);
                    ReplaceWord_Ofice2("{date}", DATA_ss, wordDocument3);
                    ReplaceWord_Ofice2("{time}", Time_s, wordDocument3);
                }

                if (checkBox6.Checked == true) // Отменить вставку Фамилия/Дата/Время
                {
                    if (checkBox5.Checked == true)
                    {
                        ReplaceWord_Ofice2("{№ИЗД2}", "", wordDocument3);
                        ReplaceWord_Ofice2("{№ИЗД1}", "ИЗД №" + textBox131.Text, wordDocument3);
                    }
                    if (checkBox5.Checked == false)
                    {
                        ReplaceWord_Ofice2("{№ИЗД1}", "ИЗД №" + textBox131.Text, wordDocument3);
                        ReplaceWord_Ofice2("{№ИЗД2}", "ИЗД №" + textBox91.Text, wordDocument3);
                    }
                    ReplaceWord_Ofice2("{family}", "", wordDocument3);
                    ReplaceWord_Ofice2("{date}", "", wordDocument3);
                    ReplaceWord_Ofice2("{time}", "",  wordDocument3);
                }

                //настройка сохранения файла
                if (textBox131.Text != "" && textBox91.Text != "")
                {
                    wordDocument3.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "IZD №" + textBox131.Text + "," + textBox91.Text + "_" + DATA_ss + ".docx");
                }
                if (textBox131.Text != "" && (textBox91.Text == "") || (checkBox5.Checked == true))
                {
                    wordDocument3.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "IZD №" + textBox131.Text + "_" + DATA_ss + ".docx");
                }
                if (textBox131.Text == "" && textBox91.Text == "")
                {
                    wordDocument3.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "IZD №" + "_" + DATA_ss + ".docx");
                }
                if (textBox131.Text == "" && textBox91.Text != "" && checkBox5.Checked == false)
                {
                    wordDocument3.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "IZD №" + textBox91.Text + "_" + DATA_ss + ".docx");
                }


                //wordDocument.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + DATA_ss + "Prot_№ " + nomer + ".docx");
                //wordDocument3.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/Proverka/" + "IZD №" + textBox131.Text + " " + textBox91.Text + "_" + DATA_ss + ".docx");
                wordApp3.Visible = true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

 

    }
}