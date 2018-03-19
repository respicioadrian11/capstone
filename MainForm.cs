


using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.Structure;
using Emgu.CV.CvEnum;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Data;

namespace MultiFaceRec
{
    public partial class FrmPrincipal : Form
    {

        //Declararation of all variables, vectors and haarcascades
        Image<Bgr, Byte> currentFrame;
        Emgu.CV.Capture webcam;
        HaarCascade face;
        //HaarCascade eye;
        MCvFont font = new MCvFont(FONT.CV_FONT_HERSHEY_TRIPLEX, 0.5d, 0.5d);
        Image<Gray, byte> result, TrainedFace = null;
        Image<Gray, byte> gray = null;
        List<Image<Gray, byte>> trainingImages = new List<Image<Gray, byte>>();
        List<string> labels = new List<string>();
        List<string> NamePersons = new List<string>();
        int ContTrain, NumLabels, t;
        string name, names = null;



        private OleDbConnection connection = new OleDbConnection();
        public DataSet Dset = new DataSet();
        public OleDbDataAdapter OleDa = new OleDbDataAdapter();
        public FrmPrincipal()
        
        {


            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\FaceRecProOV\bin\Debug\Test.accdb; Jet OleDB:Database Password=harvz";
            OleDbDataAdapter OleDa = new OleDbDataAdapter("SELECT * FROM [Test]", connection);
            //Load haarcascades for face detection
            face = new HaarCascade("haarcascade_frontalface_default.xml");
            //eye = new HaarCascade("haarcascade_eye.xml"); 
            try
            {
                //Load of previus trainned faces and labels for each image
                string Labelsinfo = File.ReadAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.xml");
                string[] Labels = Labelsinfo.Split('%');
                NumLabels = Convert.ToInt16(Labels[0]);
                ContTrain = NumLabels;
                string LoadFaces;

                for (int tf = 1; tf < NumLabels + 1; tf++)
                {
                    LoadFaces = "face" + tf + ".bmp";
                    trainingImages.Add(new Image<Gray, byte>(Application.StartupPath + "/TrainedFaces/" + LoadFaces));
                    labels.Add(Labels[tf]);
                }

            }
            catch (Exception )
            {
                //MessageBox.Show(e.ToString());
                MessageBox.Show("Nothing in binary database, please add at least a face(Simply train the prototype with the Add Face Button).", "Triained faces load", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, EventArgs e)
        {

        }
       

 

    
        private void button3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = "C:";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = DateTime.Now.ToLongDateString();
            saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx|Excel Files(2010)|*.xlsx";
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing);
                ExcelApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    ExcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;

                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                    }
                }
                ExcelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                ExcelApp.ActiveWorkbook.Saved = true;
                MessageBox.Show("Exported Successfully" + saveFileDialog1.FileName);


            }


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel3.Show();
            /*
           
            */
        }

        private void FrmPrincipal_Load(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox2.Text== "ascobal")
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = "DELETE * FROM [Table1]";
                cmd.Connection = connection;
                connection.Open();
                if (connection.State == System.Data.ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Database Reset Successfully");
                        connection.Close();
                        textBox2.Text = "";
                        panel3.Hide();

                    }
                    catch (OleDbException)
                    {

                        connection.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Connection Failed");
                    
                }

            }
            else{
                MessageBox.Show("Nd ka pede librarian lan dapat");
                textBox2.Text = "";
                panel3.Hide();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel3.Hide();
        }

        private void imageBox1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Initialize the capture device
            webcam = new Emgu.CV.Capture();
            webcam.QueryFrame();
            //Initialize the FrameGraber event
            Application.Idle += new EventHandler(Framewebcam);
            button1.Enabled = false;

        }


        private void button2_Click(object sender, System.EventArgs e)
        {

            try
            {
                //Trained face counter
                ContTrain = ContTrain+1;

                //Get a gray frame from capture device
                gray = webcam.QueryGrayFrame().Resize(320, 240, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);

                //Face Detector
                MCvAvgComp[][] facesDetected = gray.DetectHaarCascade(
                face,
                1.2,
                10,
                Emgu.CV.CvEnum.HAAR_DETECTION_TYPE.DO_CANNY_PRUNING,
                new Size(20, 20));

                //Action for each element detected
                foreach (MCvAvgComp f in facesDetected[0])
                {
                    TrainedFace = currentFrame.Copy(f.rect).Convert<Gray, byte>();
                    break;
                }

                //resize face detected image for force to compare the same size with the 
                //test image with cubic interpolation type method
                TrainedFace = result.Resize(100, 100, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);
                trainingImages.Add(TrainedFace);
                labels.Add(textBox1.Text);




                //Show face added in gray scale
                imageBox1.Image = TrainedFace;

                //Write the number of triained faces in a file text for further load
                File.WriteAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.xml", trainingImages.ToArray().Length.ToString() + "%");

                //Write the labels of triained faces in a file text for further load
                for (int i = 1; i < trainingImages.ToArray().Length + 1; i++)
                {
                    trainingImages.ToArray()[i - 1].Save(Application.StartupPath + "/TrainedFaces/face" + i + ".bmp");
                    File.AppendAllText(Application.StartupPath + "/TrainedFaces/TrainedLabels.xml", labels.ToArray()[i - 1] + "%");
                }

                MessageBox.Show(textBox1.Text + "´s face detected and added :)", "Training OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //textBox1.Text= ("");

            }
            catch
            {
                MessageBox.Show("Enable the face detection first", "Training Fail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }



        void Framewebcam(object sender, EventArgs e)
        {

            //label3.Text = "0";
            //label4.Text = "";
            NamePersons.Add("");


            //Get the current frame form capture device
            currentFrame = webcam.QueryFrame().Resize(320, 240, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);

            //Convert it to Grayscale
            gray = currentFrame.Convert<Gray, Byte>();

            //Face Detector
            MCvAvgComp[][] facesDetected = gray.DetectHaarCascade(
          face,
          1.2,
          10,
          Emgu.CV.CvEnum.HAAR_DETECTION_TYPE.DO_CANNY_PRUNING,
          new Size(20, 20));

            //Action for each element detected
            foreach (MCvAvgComp f in facesDetected[0])
            {
                t = t + 1;
                result = currentFrame.Copy(f.rect).Convert<Gray, byte>().Resize(100, 100, Emgu.CV.CvEnum.INTER.CV_INTER_CUBIC);
                //draw the face detected in the 0th (gray) channel with blue color
                currentFrame.Draw(f.rect, new Bgr(Color.Red), 2);


                if (trainingImages.ToArray().Length != 0)
                {
                    //TermCriteria for face recognition with numbers of trained images like maxIteration
                    MCvTermCriteria termCrit = new MCvTermCriteria(ContTrain, 0.001);

                    //Eigen face recognizer
#pragma warning disable CS0436 // Type conflicts with imported type
#pragma warning disable CS0436 // Type conflicts with imported type
                    EigenObjectRecognizer recognizer = new EigenObjectRecognizer(
#pragma warning restore CS0436 // Type conflicts with imported type
#pragma warning restore CS0436 // Type conflicts with imported type
                       trainingImages.ToArray(),
                       labels.ToArray(),
                       3000,
                       ref termCrit);

                    name = recognizer.Recognize(result);

                    //Draw the label for each face detected and recognized
                    currentFrame.Draw(name, ref font, new Point(f.rect.X - 2, f.rect.Y - 2), new Bgr(Color.LightGreen));

                }


                NamePersons[t - 1] = name;
                NamePersons.Add("");


                //Set the number of faces detected on the scene


                label3.Text = dataGridView1.Rows.Count.ToString();
                //facesDetected[0].Length.ToString();



                /*
                //Set the region of interest on the faces

                gray.ROI = f.rect;
                MCvAvgComp[][] eyesDetected = gray.DetectHaarCascade(
                   eye,
                   1.1,
                   10,
                   Emgu.CV.CvEnum.HAAR_DETECTION_TYPE.DO_CANNY_PRUNING,
                   new Size(20, 20));
                gray.ROI = Rectangle.Empty;

                foreach (MCvAvgComp ey in eyesDetected[0])
                {
                    Rectangle eyeRect = ey.rect;
                    eyeRect.Offset(f.rect.X, f.rect.Y);
                    currentFrame.Draw(eyeRect, new Bgr(Color.Blue), 2);
                }
                 */

            }
            t = 0;

            //Names concatenation of persons recognized
            for (int nnn = 0; nnn < facesDetected[0].Length; nnn++)
            {
                names = names + NamePersons[nnn];
            }
            //Show the faces procesed and recognized
            imageBoxFramewebcam.Image = currentFrame;
            label4.Text = names;
            names = "";
            //Clear the list(vector) of names
            NamePersons.Clear();


            if (label4.Text != names)
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = "INSERT INTO [Table1] ([sinfo],[time]) VALUES (@info, @time)";
                cmd.Connection = connection;
                connection.Open();
                if (connection.State == System.Data.ConnectionState.Open)
                {
                    cmd.Parameters.Add("@info", OleDbType.VarChar).Value = label4.Text;
                    cmd.Parameters.Add("@time", OleDbType.VarChar).Value = DateTime.Now.ToShortTimeString();
                    try
                    {
                        cmd.ExecuteNonQuery();
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = label4.Text;
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[2].Value = DateTime.Now.ToShortTimeString();
                       
                        connection.Close();

                    }
                    catch (OleDbException )
                    {
                       
                        connection.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Connection Failed");
                }
                /*
             connection.Open();
             OleDbCommand comd = new OleDbCommand();
             comd.CommandText = "SELECT * FROM Table1";
             comd.Connection = connection;
             connection.Open();
             cmd.ExecuteNonQuery();

             dataGridView1.Rows.Add();
             dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = dataGridView1.Rows.Count - 1;
             dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = label4.Text;
             dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[2].Value = DateTime.Now.ToShortTimeString();
             */
            }

        }

       }

   }
           
            

         
        
   
    


