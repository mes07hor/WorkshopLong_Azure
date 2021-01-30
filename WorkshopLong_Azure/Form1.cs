using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;


namespace WorkshopLong_Azure
{
    public partial class Form1 : Form
    {
        private SpeechRecognizer recognizer;
        TaskCompletionSource<int> stopRecognition;
        public Form1()
        {
            InitializeComponent();
            //var config = SpeechConfig.FromSubscription("*******************", "japaneast");
            var config = SpeechConfig.FromSubscription("********************", "japaneast");

            config.EnableDictation();
            config.SpeechRecognitionLanguage = "ja-JP";
            var audioConfig = AudioConfig.FromDefaultMicrophoneInput();
            recognizer = new SpeechRecognizer(config,audioConfig);
            stopRecognition = new TaskCompletionSource<int>();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            //GetTableCatalog();
            //comboBox1.DataSource = TableCatalog;
            comboBox1.Items.Add("Adults1");
            comboBox1.Items.Add("Adults2");
            comboBox1.Items.Add("Students1");
            comboBox1.Items.Add("Students2");
            comboBox1.Items.Add("voicerecord");
        }
        private static double time = 0; //total speaking time


        private static string name = null;    //participant name Initialize
        public bool nameInput = false;
        static bool speechbool = false;

        
        private async void startRecognition(object sender, EventArgs e)
        {
            name = textBox1.Text;

            nameInput = true;


            System.Diagnostics.Debug.WriteLine("Say something");


            // Starts speech recognition, and returns after a single utterance is recognized. The end of a
            // single utterance is determined by listening for silence at the end or until a maximum of 15
            // seconds of audio is processed.  The task returns the recognition text as result. 
            // Note: Since RecognizeOnceAsync() returns only a single utterance, it is suitable only for single
            // shot recognition like command or query. 
            // For long-running multi-utterance recognition, use StartContinuousRecognitionAsync() instead.
            //var result = await recognizer.RecognizeOnceAsync();

            recognizer.Recognizing += (s, e) =>
            {
                System.Diagnostics.Debug.WriteLine($"RECOGNIZING: Text={e.Result.Text}");
                label5.Invoke(new Action(()=> { label5.Text = "認識中"; }));

            };

            recognizer.Recognized += (s, e) =>
            {
                if (e.Result.Reason == ResultReason.RecognizedSpeech)
                {
                    System.Diagnostics.Debug.WriteLine($"RECOGNIZED: Text={e.Result.Text}");
                    //string text1 = e.Result.Text;
                    //label2.Text = text1;
                    label1.Invoke(new Action(() => { label1.Text = e.Result.Text; }));

                    System.Diagnostics.Debug.WriteLine($"RECOGNIZED: Time={e.Result.Duration}");
                    time = time + e.Result.Duration.TotalSeconds;
                    System.Diagnostics.Debug.WriteLine($"RECOGNIZED: Total Time={time}");
                    label5.Invoke(new Action(()=> { label5.Text = "待機中"; }));

                    submitsentence(e.Result.Text);

                }
                else if (e.Result.Reason == ResultReason.NoMatch)
                {
                    System.Diagnostics.Debug.WriteLine($"NOMATCH: Speech could not be recognized.");
                }
                speechbool = true;
            };



            recognizer.Canceled += (s, e) =>
            {
                Console.WriteLine($"CANCELED: Reason={e.Reason}");

                if (e.Reason == CancellationReason.Error)
                {
                    Console.WriteLine($"CANCELED: ErrorCode={e.ErrorCode}");
                    Console.WriteLine($"CANCELED: ErrorDetails={e.ErrorDetails}");
                    Console.WriteLine($"CANCELED: Did you update the subscription info?");
                }

                stopRecognition.TrySetResult(0);
            };

            recognizer.SessionStopped += (s, e) =>
            {
                Console.WriteLine("\n    Session stopped event.");
                stopRecognition.TrySetResult(0);
            };

            await recognizer.StartContinuousRecognitionAsync();
            await Task.WhenAny(stopRecognition.Task);
        }
        static bool btnClickFirstTime = true;
        private void button1_Click(object sender, EventArgs e)
        {
            if (btnClickFirstTime)
            {
                startRecognition(sender, e);
                btnClickFirstTime = false;
                label2.Text="ようこそ、"+textBox1.Text.ToString()+"さん";
            }

        }
        
        static void submitsentence(string sentence)
        {
            string constr = @"Data Source=***************.rds.amazonaws.com,1433;
                                Initial Catalog=Workshop;Connect Timeout=20;Persist Security Info=True;User ID=*****;Password=*******";



            using (SqlConnection con = new SqlConnection(constr))
            using (var command = con.CreateCommand())
            {
                try
                {
                    DateTime date = DateTime.Now;
                    string Date = date.ToString("yyyy-MM-dd HH:mm:ss.fff");         

                    string commandText = "INSERT INTO " + "SentenceTest2" + " (Date, sentence, username,selectedgroup) VALUES(@Date, @sentence, @username,@selectedgroup)";
                    con.Open();
                    //command.CommandText = @"INSERT INTO voicerecord (Date,time,username) VALUES (@Date,@time,@username)";
                    command.CommandText = @commandText;

                    command.Parameters.Add(new SqlParameter("@Date", Date));
                    command.Parameters.Add(new SqlParameter("@sentence", sentence));
                    command.Parameters.Add(new SqlParameter("@username", name));
                    command.Parameters.Add(new SqlParameter("@selectedgroup", selectedTable));


                    command.ExecuteNonQuery();
                    Console.WriteLine("{0} success", Date);

                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                    throw;
                }
                finally
                {
                    con.Close();
                }
            }
        }

        static void submitdata(double time)
        {
            
            
            string constr = @"Data Source=*************.rds.amazonaws.com,1433;
                                Initial Catalog=Workshop;Connect Timeout=20;Persist Security Info=True;User ID=*****;Password=*****";



            using (SqlConnection con = new SqlConnection(constr))
            using (var command = con.CreateCommand())
            {
                try
                {
                    DateTime date = DateTime.Now;
                    string Date = date.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    string commandText;

                    switch (selectedTable)
                    {
                        case "Adults1":
                        case "Adults2":
                            commandText = "INSERT INTO " + "Adults" + " (Date, time, username,selectedgroup) VALUES(@Date, @time, @username,@selectedgroup)";
                            break;
                        case "Students1":
                        case "Students2":
                            commandText = "INSERT INTO " + "Students" + " (Date, time, username,selectedgroup) VALUES(@Date, @time, @username,@selectedgroup)";
                            break;
                        default:
                            commandText = "INSERT INTO " + "voicerecord" + " (Date, time, username,selectedgroup) VALUES(@Date, @time, @username,@selectedgroup)";
                            break;
                    }
                    con.Open();
                    //command.CommandText = @"INSERT INTO voicerecord (Date,time,username) VALUES (@Date,@time,@username)";
                    command.CommandText = @commandText;

                    command.Parameters.Add(new SqlParameter("@Date", Date));
                    command.Parameters.Add(new SqlParameter("@time", time));
                    command.Parameters.Add(new SqlParameter("@username", name));
                    command.Parameters.Add(new SqlParameter("@selectedgroup", selectedTable));

                    command.ExecuteNonQuery();
                    Console.WriteLine("{0} success", Date);

                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                    throw;
                }
                finally
                {
                    con.Close();
                }
            }

        }
        private static string selectedTable;


        //SQL to get table catalog
        private static readonly string SelectTableListSql = "SELECT * FROM sys.objects where type='U'";

        //list of table catalog
        static string[] TableCatalog = new string[10];
        static List<string> intermediateTableCatalog = new List<string>();

        static void GetTableCatalog()
        {
            string constr = @"Data Source=*********.rds.amazonaws.com,1433;
                                Initial Catalog=Workshop;Connect Timeout=20;Persist Security Info=True;User ID=******;Password=********";
            // コネクションを生成します。
            using (var connection = new SqlConnection(constr))
                try
                {
                    // コマンドオブジェクトを作成します。
                    using (var command = connection.CreateCommand())
                    {
                        // コネクションをオープンします。
                        connection.Open();

                        // テーブル一覧取得のSQLを実行します。
                        command.CommandText = SelectTableListSql;
                        var reader = command.ExecuteReader();
                        // 取得結果を確認します。
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                var name = reader["name"] as string;
                                Console.WriteLine(name);
                                intermediateTableCatalog.Add(name);
                            }
                            TableCatalog = intermediateTableCatalog.ToArray();
                        }
                        for (int i = 0; i < 5; i++)
                        {
                            Console.WriteLine(TableCatalog[i]);
                        }
                    }
                }

                // 例外が発生した場合
                catch (Exception e)
                {
                    // 例外の内容を表示します。
                    Console.WriteLine(e.Message);
                }
                finally
                {
                    connection.Close();
                }
        }

        static void timer() //submit data every minitute
        {
            DateTime intervaltime = DateTime.Now;
            int nowMinnite = intervaltime.Minute;
            int nowSecond = intervaltime.Second;
            if (nowMinnite % 1 == 0 && nowSecond == 00)   //every minute
            {

                submitdata(time);
                System.Diagnostics.Debug.WriteLine("submmit");
                time = 0;         //reset total speaking time
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (nameInput)
            {
                timer();
            }
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int index = comboBox1.SelectedIndex;
            selectedTable = comboBox1.Items[index].ToString();
            System.Diagnostics.Debug.WriteLine("selected table is", selectedTable);
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            //Task.WaitAny(new[] { stopRecognition.Task });
            //startRecognition(sender,e);
            
        }
    }
}
