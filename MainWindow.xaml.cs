using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.Reflection;
using System.Security.Cryptography;
using System.Collections;

namespace Lab5
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private int counter = 0;
        public MainWindow()
        {
            InitializeComponent();
        }

        private ISldWorks app = null;

        //открыть чертёж
        ModelDoc2 doc;

        private void button_Click(object sender, RoutedEventArgs e)
        {
            

            //массивы размеров детали
            var s = new Dictionary<string, double>() {
                { "L1", 0.012 },
                { "L2", 0.050 },
                { "L3", 0.095 },
                { "L4", 0.120 },
                { "L5", 0.090 },
                { "L6", 0.050 },
                { "L7", 0.025 },
                { "L8", 0.025 },
                { "L9", 0.090 },
                { "L10", 0.050 },
                { "L11", 0.020 }
            };

            double x0 = (s["L5"] - s["L2"]) / 2, y0 = s["L3"], z0 = 0;

            string msg = "";

            x0 /= 1000; y0 /= 1000; z0 /= 1000;

            if (msg != "")
            {
                MessageBox.Show(msg, "Ошибка");
                button.Content = "Нарисовать";
                return;
            }

            //открыть SolidWorks либо получить экземпляр открытого приложения.
            try
            {
                //попытка открыть либо получить открытый через SolidWorks Api
                app = new SldWorks();

                //задаём размер окна - на весь экран
                app.FrameState = (int)swWindowState_e.swWindowMaximized;

                //делаем окно видимым
                app.Visible = true;
            }
            catch
            {
                try
                {
                    //если перый вариант не сработал - пробуем получить открытый SolidWorks
                    app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch
                {
                    //если и это не помогло, тогда что-то не так
                    MessageBox.Show("Не удалось открыть SolidWorks либо найти открытое приложение.");
                    return;
                }
            }



            //если нет открытого чертежа
            if (app.ActiveDoc == null)
            {
                //создать и открыть
                doc = (ModelDoc2)app.INewDrawing((int)swDwgTemplates_e.swDwgTemplateCustom);

                //задать размеры в миллиметрах
                doc.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
            }

            //size_x и  size_y - отступы размеров от чертежа
            double size_x = 0.01, size_y = 0.01, size_z = 0.01, x = x0, y = y0, z = z0, x1 = x0, y1 = y0, z1 = z0;

            //получаем открытый документ
            doc = (ModelDoc2)app.ActiveDoc;

            //получем ссылку на интерфейс, ответственный за рисование
            SketchManager sm = doc.SketchManager;

            //переменная, обозначающая ограничение скругление
            var tangent = "sgTANGENT";

            //выбираем какое всплывающее окно отключить
            int pref_toggle = (int)swUserPreferenceToggle_e.swInputDimValOnCreate;

            //отключаем всплыващее окно нанесения размеров
            app.SetUserPreferenceToggle(pref_toggle, false);

            var sqrtFrom2 = Math.Sqrt(2);
            var sqrtFrom125 = Math.Sqrt(125) / 1000;
            var sqrtFrom525 = Math.Sqrt(525) / 1000;

            doc.ClearSelection();

            sm.InsertSketch(false);


            var line01 = sm.CreateLine(x, y, z, x, y - s["L2"], z);
            doc.ClearSelection();

            var line02 = sm.CreateLine(x, y - s["L2"], z, x + s["L2"], y - s["L2"], z);
            doc.ClearSelection();

            var line03 = sm.CreateLine(x + s["L2"], y - s["L2"], z, x + s["L2"], y, z);
            doc.ClearSelection();



            x += s["L2"];
            x1 = x + (s["L5"] - s["L2"]) / 2; y1 = y0; z1 = z0;
            var line0 = sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();

            x += (s["L5"] - s["L2"]) / 2;
            y1 -= s["L3"];
            var line1 = sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();


            y = y1;
            x1 -= s["L5"];
            var line2 = sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();



            x = x1;
            y1 = y0;
            var line3 = sm.CreateLine(x, y, z0, x1, y1, z0);
            doc.ClearSelection();

            y = y1;
            x1 = x0;
            var line4 = sm.CreateLine(x, y, z, x1, y1, z0);
            doc.ClearSelection();


            var feature = featureExtrusion(s["L4"], false);

            var faces8 = feature.GetFaces();
            var ent8 = faces8[6] as Entity;
            ent8.Select(true);

            sm.InsertSketch(false);

            x = -s["L1"];
            x1 = -s["L4"];
            y = 0;
            y1 = y;

            sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();

            x = x1;
            y1 -= s["L3"] - s["L8"];

            sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();

            y = y1;
            x1 += s["L9"];

            sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();

            x = x1;
            y1 += s["L10"];

            sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();


            y = y1;
            x1 = -s["L1"];

            sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();

            x = x1;
            y = y1;
            y1 = 0;
            x1 = -s["L1"];

            sm.CreateLine(x, y, z, x1, y1, z1);
            doc.ClearSelection();

            featureCut(s["L5"]);
            doc.ClearSelection();

            faces8 = feature.GetFaces();
            ent8 = faces8[7] as Entity;
            ent8.Select(true);


            sm.InsertSketch(false);

            sm.CreateCircle(0.025, 0.075, 0, 0.045, 0.105, 0);
            doc.ClearSelection();

            featureCut(0.025);
            doc.ClearSelection();

            faces8 = feature.GetFaces();
            ent8 = faces8[7] as Entity;
            ent8.Select(true);


            sm.InsertSketch(false);


            sm.CreateLine(-0.030, 0.075, 0, -0.030, 0.130, 0);
            doc.ClearSelection();


            sm.CreateArc(0.025, 0.075, 0, -0.020, 0.075, 0, 0.025, 0.120, 0, -1);
            doc.ClearSelection();

            sm.CreateLine(-0.020, 0.075, 0, -0.030, 0.075, 0);
            doc.ClearSelection();

            sm.CreateLine(-0.030, 0.130, 0, 0.025, 0.120, 0);

            featureCut(0.025);
            doc.ClearSelection();

            faces8 = feature.GetFaces();


            ent8 = faces8[0] as Entity;
            ent8.Select(true);


            sm.InsertSketch(false);
            sm.CreateLine(0.025, 0.120, 0, 0.070, 0.120, 0);
            doc.ClearSelection();


            sm.CreateArc(0.025, 0.075, 0, 0.025, 0.120, 0, 0.070, 0.075, 0, -1);
            doc.ClearSelection();

            sm.CreateLine(0.070, 0.075, 0, 0.075, 0.075, 0);
            doc.ClearSelection();

            sm.CreateLine(0.070, 0.120, 0, 0.075, 0.075, 0);

            featureCut(0.025);
            doc.ClearSelection();


            //включаем окно ввода значения размера (выключали перед построением)
            app.SetUserPreferenceToggle(pref_toggle, true);

            //меняем "обратно" название кнопки, т.к. все построения завершены
            button.Content = "Start Paint";

            //обнуляем ссылки на документ
            doc = null;
        }

        /// <summary>
        /// Вытянуть бобышку
        /// </summary>
        /// <param name="deepth">высота выдавливания</param>
        /// <param name="dir">направление выдвливания</param>
        /// <returns>объект бобышка</returns>
        private Feature featureExtrusion(double deepth, bool dir)
        {
            return doc.FeatureManager.FeatureExtrusion2(true, false, dir,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                deepth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        private void selectPlane(string name, string obj = "PLANE")
        {
            //выбрать плоскость по имени
            doc.Extension.SelectByID2(name, obj, 0, 0, 0, false, 0, null, 0);
        }

        /// <summary>
        /// Вырезать по контуру
        /// </summary>
        /// <param name="deepth">глубина выреза</param>
        /// <param name="flip">вырезать внутри контура или снаружи</param>
        /// <param name="mode">режим выреза</param>
        /// <returns>объект "вырез"</returns>
        private Feature featureCut(double deepth, bool flip = false, swEndConditions_e mode = swEndConditions_e.swEndCondBlind)
        {
            return doc.FeatureManager.FeatureCut2(true, flip, false, (int)mode, (int)mode,
                deepth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private void MainWindow1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //закрыть SolidWorks и закрыть все документы в т.ч. несохраненные
            app.CloseAllDocuments(true);
            app.ExitApp();
            app = null;

            //явно закрываем swVBAServers
            try
            {
                Process[] processes = Process.GetProcessesByName("swvbaserver");
                foreach (var process in processes)
                    process.Kill();
            }
            catch { }
        }
        int step = 1;


        double x0 = (0.090 - 0.050) / 2, y0 = 0.095, z0 = 0, x = 0, y = 0, z = 0, x1 = 0, y1 = 0, z1 = 0;

        string msg = "";

        Feature feature =null;

        double size_x = 0.01, size_y = 0.01, size_z = 0.01;
        SketchManager sm = null;
        private void button_Click_2(object sender, RoutedEventArgs e)
        {


            //массивы размеров детали
            var s = new Dictionary<string, double>() {
                { "L1", 0.012 },
                { "L2", 0.050 },
                { "L3", 0.095 },
                { "L4", 0.120 },
                { "L5", 0.090 },
                { "L6", 0.050 },
                { "L7", 0.025 },
                { "L8", 0.025 },
                { "L9", 0.090 },
                { "L10", 0.050 },
                { "L11", 0.020 }
            };
            switch (step)
            {
                case 1:
                    if (msg != "")
                    {
                        MessageBox.Show(msg, "Ошибка");
                        button.Content = "Нарисовать";
                        return;
                    }

                    //открыть SolidWorks либо получить экземпляр открытого приложения.
                    try
                    {
                        //попытка открыть либо получить открытый через SolidWorks Api
                        app = new SldWorks();

                        //задаём размер окна - на весь экран
                        app.FrameState = (int)swWindowState_e.swWindowMaximized;

                        //делаем окно видимым
                        app.Visible = true;
                    }
                    catch
                    {
                        try
                        {
                            //если перый вариант не сработал - пробуем получить открытый SolidWorks
                            app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                        }
                        catch
                        {
                            //если и это не помогло, тогда что-то не так
                            MessageBox.Show("Не удалось открыть SolidWorks либо найти открытое приложение.");
                            return;
                        }
                    }



                    //если нет открытого чертежа
                    if (app.ActiveDoc == null)
                    {
                        //создать и открыть
                        doc = (ModelDoc2)app.INewDrawing((int)swDwgTemplates_e.swDwgTemplateCustom);

                        //задать размеры в миллиметрах
                        doc.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
                    }

                    //size_x и  size_y - отступы размеров от чертежа
                    size_x = 0.01; size_y = 0.01; size_z = 0.01; x = x0; y = y0; z = z0; x1 = x0; y1 = y0; z1 = z0;

                    //получаем открытый документ
                    doc = (ModelDoc2)app.ActiveDoc;

                    //получем ссылку на интерфейс, ответственный за рисование
                    sm = doc.SketchManager;

                    //переменная, обозначающая ограничение скругление
                    var tangent = "sgTANGENT";

                    //выбираем какое всплывающее окно отключить
                    int pref_toggle = (int)swUserPreferenceToggle_e.swInputDimValOnCreate;

                    //отключаем всплыващее окно нанесения размеров
                    app.SetUserPreferenceToggle(pref_toggle, false);

                    var sqrtFrom2 = Math.Sqrt(2);
                    var sqrtFrom125 = Math.Sqrt(125) / 1000;
                    var sqrtFrom525 = Math.Sqrt(525) / 1000;

                    doc.ClearSelection();

                    sm.InsertSketch(false);

                    button.Content = "Processing";
                    button.Dispatcher.Invoke(DispatcherPriority.Render, new Action(delegate () { }));

                    string front = "Front Plane", top = "Top Plane", right = "Right Plane";
                    x0 /= 1000; y0 /= 1000; z0 /= 1000;
                    x = x0;
                    y = y0;
                    z = z0;
                    x1 = x0;
                    y1 = y0;
                    z1 = z0;

                    

               

                  
                    var line01 = sm.CreateLine(x, y, z, x, y - s["L2"], z);
                    doc.ClearSelection();

                    var line02 = sm.CreateLine(x, y - s["L2"], z, x + s["L2"], y - s["L2"], z);
                    doc.ClearSelection();

                    var line03 = sm.CreateLine(x + s["L2"], y - s["L2"], z, x + s["L2"], y, z);
                    doc.ClearSelection();
                    break;
                case 3:
                    x += s["L2"];
                    x1 = x + (s["L5"] - s["L2"]) / 2; y1 = y0; z1 = z0;
                    var line0 = sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();

                    x += (s["L5"] - s["L2"]) / 2;
                    y1 -= s["L3"];
                    var line1 = sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();
                    break;
                case 4:
                    y = y1;
                    x1 -= s["L5"];
                    var line2 = sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();



                    x = x1;
                    y1 = y0;
                    var line3 = sm.CreateLine(x, y, z0, x1, y1, z0);
                    doc.ClearSelection();
                    break;
                case 5:
                    y = y1;
                    x1 = x0;
                    var line4 = sm.CreateLine(x, y, z, x1, y1, z0);
                    doc.ClearSelection();
                    break;
                case 6:
                    feature = featureExtrusion(s["L4"], false);
                    var faces8 = feature.GetFaces();
                    var ent8 = faces8[6] as Entity;
                    ent8.Select(true);

                    sm.InsertSketch(false);

                    x = -s["L1"];
                    x1 = -s["L4"];
                    y = 0;
                    y1 = y;

                    sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();
                    break;
                case 7:
                    x = x1;
                    y1 -= s["L3"] - s["L8"];

                    sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();

                    y = y1;
                    x1 += s["L9"];

                    sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();

                    x = x1;
                    y1 += s["L10"];

                    sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();
                    break;
                case 8:
                    y = y1;
                    x1 = -s["L1"];

                    sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();

                    x = x1;
                    y = y1;
                    y1 = 0;
                    x1 = -s["L1"];

                    sm.CreateLine(x, y, z, x1, y1, z1);
                    doc.ClearSelection();

                    featureCut(s["L5"]);
                    doc.ClearSelection();
                    break;
                case 9:
                    faces8 = feature.GetFaces();
                    ent8 = faces8[7] as Entity;
                    ent8.Select(true);


                    sm.InsertSketch(false);

                    sm.CreateCircle(0.025, 0.075, 0, 0.045, 0.105, 0);
                    doc.ClearSelection();

                    featureCut(0.025);
                    doc.ClearSelection();
                    break;
                case 10:
                    faces8 = feature.GetFaces();
                    ent8 = faces8[7] as Entity;
                    ent8.Select(true);


                    sm.InsertSketch(false);


                    sm.CreateLine(-0.030, 0.075, 0, -0.030, 0.130, 0);
                    doc.ClearSelection();
                    break;
                case 11:
                    sm.CreateArc(0.025, 0.075, 0, -0.020, 0.075, 0, 0.025, 0.120, 0, -1);
                    doc.ClearSelection();

                    sm.CreateLine(-0.020, 0.075, 0, -0.030, 0.075, 0);
                    doc.ClearSelection();
                    break;
                case 12:
                    sm.CreateLine(-0.030, 0.130, 0, 0.025, 0.120, 0);

                    featureCut(0.025);
                    doc.ClearSelection();

                    faces8 = feature.GetFaces();


                    ent8 = faces8[0] as Entity;
                    ent8.Select(true);


                    sm.InsertSketch(false);
                    sm.CreateLine(0.025, 0.120, 0, 0.070, 0.120, 0);
                    doc.ClearSelection();
                    break;
                case 13:
                    sm.CreateArc(0.025, 0.075, 0, 0.025, 0.120, 0, 0.070, 0.075, 0, -1);
                    doc.ClearSelection();

                    sm.CreateLine(0.070, 0.075, 0, 0.075, 0.075, 0);
                    doc.ClearSelection();
                    break;
                case 14:
                    sm.CreateLine(0.070, 0.120, 0, 0.075, 0.075, 0);

                    featureCut(0.025);
                    doc.ClearSelection();
                    break;
                case 16: default: break;
            }
            step++;
        }
    }
}
