﻿using Excel_XLSX_reader_writer.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Excel_XLSX_reader_writer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void RunBTN_Click(object sender, RoutedEventArgs e)
        {
            // File open dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document";
            //temp default path
            dlg.InitialDirectory = @"C:\TEMP_SK\excel";

            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                excelRead.ReadLargeFile(dlg.FileName);
                // program completion
                MessageBox.Show("Finished");
            }
        }
    }
}
