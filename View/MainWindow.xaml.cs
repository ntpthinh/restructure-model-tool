using Microsoft.ML;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Resources;

namespace RestructureModelTool.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //For testing purpose

        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private DateTime x = new DateTime(2020, 4, 10);
        private string folderPath = "";
        MLContext mlContext;
        public MainWindow()
        {
            InitializeComponent();
            mlContext = new MLContext();
        }




        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("in");
            DataViewSchema modelSchema;
            // Load trained model
            MessageBox.Show("out");
            try
            {
                Uri uri = new Uri("/RestructureModelTool;Component/model.zip", UriKind.Relative);
                StreamResourceInfo info = Application.GetResourceStream(uri);
                MessageBox.Show("stream");
                ITransformer trainedModel = mlContext.Model.Load(info.Stream, out modelSchema);
                MessageBox.Show("load model");
                ClassifySingleImage(mlContext, trainedModel);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show(ex.InnerException.Message);
                MessageBox.Show(ex.InnerException.InnerException.Message);
                MessageBox.Show(ex.InnerException.InnerException.StackTrace);
            }


        }
        public static void ClassifySingleImage(MLContext mlContext, ITransformer model)
        {
            string _predictSingleImage = @"D:\Images\sofa3.jpg";

            var imageData = new ImageData()
            {
                ImagePath = _predictSingleImage
            };
            // Make prediction function (input = ImageData, output = ImagePrediction)
            var predictor = mlContext.Model.CreatePredictionEngine<ImageData, ImagePrediction>(model);
            MessageBox.Show("predictionengine");

            var prediction = predictor.Predict(imageData);

            MessageBox.Show($"Image: {Path.GetFileName(imageData.ImagePath)} predicted as: {prediction.PredictedLabelValue} with score: {prediction.Score.Max()} ");

        }
        public class ImageData
        {
            public string ImagePath;

            public string Label;
        }
        public class ImagePrediction : ImageData
        {
            public float[] Score;

            public string PredictedLabelValue;
        }
    }
}
