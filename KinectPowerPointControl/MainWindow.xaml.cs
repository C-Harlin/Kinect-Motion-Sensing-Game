using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using Microsoft.Speech.Recognition;
using System.Threading;
using System.IO;
using Microsoft.Speech.AudioFormat;
using System.Diagnostics;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;

namespace KinectPowerPointControl
{
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll", EntryPoint = "keybd_event", SetLastError = true)]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        /// <summary>
        /// kinect传感器对象
        /// </summary>
        KinectSensor sensor;

        /// <summary>
        /// 语音识别对象
        /// </summary>
        SpeechRecognitionEngine speechRecognizer;

        /// <summary>
        /// 计时器
        /// </summary>
        DispatcherTimer readyTimer;

        /// <summary>
        /// 颜色数据对象
        /// </summary>
        byte[] colorBytes;

        /// <summary>
        /// 骨骼数据对象
        /// </summary>
        Skeleton[] skeletons;

        /// <summary>
        /// 是否显示左上角圆形图案对象
        /// </summary>
        bool isCirclesVisible = true;

        /// <summary>
        /// forword状态变量
        /// </summary>
        bool isForwardGestureActive = false;

        /// <summary>
        /// back状态变量
        /// </summary>
        bool isBackGestureActive = false;

        //后退状态变量
        bool myBackFlag = false;

        //前进状态变量
        bool myForwardFlag = false;
        //跳跃状态变量
        bool myJumpFlag = false;
        //拳击状态变量
        bool myBoxingFlag = false;
        //角色方位变换状态变量
        bool roleOrienFlag = false;
        // 飞踢
        bool myFlyKick = false;


        //是否双人游戏
        bool dualFlag = false;

        //角色2标志
        bool isForwardGestureActive2 = false;
        bool isBackGestureActive2 = false;
        bool myBackFlag2 = false;
        bool myForwardFlag2 = false;
        bool myJumpFlag2 = false;
        bool myBoxingFlag2 = false;
        bool roleOrienFlag2 = false;
        bool myFlyKick2 = false;

        //是否是第一帧标志
        //        bool firstFrameFlag = true;

        //上一帧头部信息
        float last_head_y = 0;
        float last_head_y2 = 0;

        //30帧计时器
        int Framesecond = 0;

        /// <summary>
        /// 屏幕是否变黑标志
        /// </summary>
        bool isBlackScreenActive = false;

        /// <summary>
        /// PPT 放映标志
        /// </summary>
        bool isPresent = false;

        /// <summary>
        /// 手臂水平伸展的阈值
        /// </summary>
        private const double ArmStretchedThreshold = 0.65;

        /// <summary>
        /// 手臂垂直上举的阈值
        /// </summary>
        private const double ArmRaisedThreshold = 0.6;

        /// <summary>
        /// 头离双手距离的阈值
        /// </summary>
        private const double DistanceThreshold = 0.05;

        /// <summary>
        /// 动作被激活时的颜色笔刷
        /// </summary>
        SolidColorBrush activeBrush = new SolidColorBrush(Colors.Green);

        /// <summary>
        /// 动作未被激活的颜色笔刷
        /// </summary>
        SolidColorBrush inactiveBrush = new SolidColorBrush(Colors.Red);

        public MainWindow()
        {
            InitializeComponent();


            //当窗体被打开，运行初始化方法，窗体关闭时，去初始化 
            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);


            //Handle the content obtained from the video camera, once received. 
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
        }

        /// <summary>
        /// 开启传感器，获取数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            sensor = KinectSensor.KinectSensors.FirstOrDefault();
            //2017/12/21 by czl
            
            if (sensor == null)
            {
                MessageBox.Show("This application requires a Kinect sensor.");
                this.Close();
            }

            sensor.Start();
            //2017/12/21 by czl

            //SpeechSynthesizer voice = new SpeechSynthesizer();   //创建语音实例
            //voice.Rate = 1; //设置语速,[-10,10]
            //voice.Volume = 100; //设置音量,[0,100]
            //voice.SpeakAsync("Device connected");  //播放指定的字符串
            
            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            sensor.ColorFrameReady += new EventHandler<ColorImageFrameReadyEventArgs>(sensor_ColorFrameReady);

            sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);

            //2017/10/21 by zzl
            var parameters = new TransformSmoothParameters
            {
                Smoothing = 0.5f,
                Correction = 0.5f,
                Prediction = 0.5f,
                JitterRadius = 0.05f,
                MaxDeviationRadius = 0.04f
            };
            //2017/12/21 by czl
            /*
             voice.Rate = 1; //设置语速,[-10,10]
             voice.Volume = 100; //设置音量,[0,100]
             voice.SpeakAsync("Thank you for playing this game");

             voice.Rate = 1; //设置语速,[-10,10]
             voice.Volume = 100; //设置音量,[0,100]
             voice.SpeakAsync("Please select the game mode");
             */
            //2017/11/29 by zzl
            HideCircles();

            sensor.SkeletonStream.Enable(parameters);

            //    sensor.SkeletonStream.Enable();
            sensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(sensor_SkeletonFrameReady);

            //sensor.ElevationAngle = 10;

            Application.Current.Exit += new ExitEventHandler(Current_Exit);

            InitializeSpeechRecognition();
        }

        /// <summary>
        /// 应用退出事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Current_Exit(object sender, ExitEventArgs e)
        {
            if (speechRecognizer != null)
            {
                speechRecognizer.RecognizeAsyncCancel();
                speechRecognizer.RecognizeAsyncStop();
            }
            if (sensor != null)
            {
                sensor.AudioSource.Stop();
                sensor.Stop();
                sensor.Dispose();
                sensor = null;
            }
        }

        /// <summary>
        /// 键盘监听事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C)
            {
                //点击键盘的C键，左上角的圆形图案消失
                ToggleCircles();
            }
        }

        /// <summary>
        /// 彩色摄像头事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void sensor_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (var image = e.OpenColorImageFrame())
            {
                if (image == null)
                    return;

                if (colorBytes == null ||
                    colorBytes.Length != image.PixelDataLength)
                {
                    colorBytes = new byte[image.PixelDataLength];
                }

                image.CopyPixelDataTo(colorBytes);

                //You could use PixelFormats.Bgr32 below to ignore the alpha,
                //or if you need to set the alpha you would loop through the bytes 
                //as in this loop below
                int length = colorBytes.Length;
                for (int i = 0; i < length; i += 4)
                {
                    colorBytes[i + 3] = 255;
                }

                BitmapSource source = BitmapSource.Create(image.Width,
                    image.Height,
                    96,
                    96,
                    PixelFormats.Bgra32,
                    null,
                    colorBytes,
                    image.Width * image.BytesPerPixel);
                videoImage.Source = source;
            }
        }



        /// <summary>
        /// 骨骼事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void sensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            //clock on
            DateTime beforDT = System.DateTime.Now;

            if (Framesecond < 30)
                Framesecond = Framesecond + 1;
            else
                Framesecond = 0;

            using (var skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame == null)
                    return;

                if (skeletons == null ||
                    skeletons.Length != skeletonFrame.SkeletonArrayLength)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                }

                skeletonFrame.CopySkeletonDataTo(skeletons);
            }

            Skeleton closestSkeleton = skeletons.Where(s => s.TrackingState == SkeletonTrackingState.Tracked)
                                                .OrderBy(s => s.Position.Z * Math.Abs(s.Position.X))
                                                .FirstOrDefault();
            Skeleton Skeleton2 = skeletons.Where(s => s.TrackingState == SkeletonTrackingState.Tracked)
                                    .OrderBy(s => s.Position.Z * Math.Abs(s.Position.X))
                                    .LastOrDefault();


            if (closestSkeleton == null)
                return;

            var head = closestSkeleton.Joints[JointType.Head];
            var rightHand = closestSkeleton.Joints[JointType.HandRight];
            var leftHand = closestSkeleton.Joints[JointType.HandLeft];
            //2017/10/21 by zzl
            var leftElbow = closestSkeleton.Joints[JointType.ElbowLeft];
            var rightElbow = closestSkeleton.Joints[JointType.ElbowRight];
            var hipCenter = closestSkeleton.Joints[JointType.HipCenter];
            var leftKnee = closestSkeleton.Joints[JointType.KneeLeft];
            var rightKnee = closestSkeleton.Joints[JointType.KneeRight];
            var leftFoot = closestSkeleton.Joints[JointType.FootLeft];
            var rightFoot = closestSkeleton.Joints[JointType.FootRight];

            var head2 = Skeleton2.Joints[JointType.Head];
            var rightHand2 = Skeleton2.Joints[JointType.HandRight];
            var leftHand2 = Skeleton2.Joints[JointType.HandLeft];
            var leftElbow2 = Skeleton2.Joints[JointType.ElbowLeft];
            var rightElbow2 = Skeleton2.Joints[JointType.ElbowRight];
            var hipCenter2 = Skeleton2.Joints[JointType.HipCenter];
            var leftKnee2 = Skeleton2.Joints[JointType.KneeLeft];
            var rightKnee2 = Skeleton2.Joints[JointType.KneeRight];
            var leftFoot2 = Skeleton2.Joints[JointType.FootLeft];
            var rightFoot2 = Skeleton2.Joints[JointType.FootRight];


            if (head.TrackingState == JointTrackingState.NotTracked ||
                rightHand.TrackingState == JointTrackingState.NotTracked ||
                leftHand.TrackingState == JointTrackingState.NotTracked
                // 2017/10/21 by zzl
                || leftElbow.TrackingState == JointTrackingState.NotTracked
                || rightElbow.TrackingState == JointTrackingState.NotTracked
                || hipCenter.TrackingState == JointTrackingState.NotTracked
                || leftKnee.TrackingState == JointTrackingState.NotTracked
                || rightKnee.TrackingState == JointTrackingState.NotTracked

                || head2.TrackingState == JointTrackingState.NotTracked
                || rightHand2.TrackingState == JointTrackingState.NotTracked ||
                leftHand2.TrackingState == JointTrackingState.NotTracked
                || leftElbow2.TrackingState == JointTrackingState.NotTracked
                || rightElbow2.TrackingState == JointTrackingState.NotTracked
                || hipCenter2.TrackingState == JointTrackingState.NotTracked
                || leftKnee2.TrackingState == JointTrackingState.NotTracked
                || rightKnee2.TrackingState == JointTrackingState.NotTracked
                //|| leftFoot.TrackingState == JointTrackingState.NotTracked
                //|| rightFoot.TrackingState == JointTrackingState.NotTracked
                )
            {
                //Don't have a good read on the joints so we cannot process gestures
                return;
            }

            //调用填充头和双手位置图案的的方法
            SetEllipsePosition(ellipseHead, head, false);
            SetEllipsePosition(ellipseLeftHand, leftHand, isBackGestureActive);
            SetEllipsePosition(ellipseRightHand, rightHand, isForwardGestureActive);
            // 2017/10/21 by zzl
            //     SetEllipsePosition(ellipseLeftHand, leftHand, false);
            //     SetEllipsePosition(ellipseRightHand, rightHand, false);
            SetEllipsePosition(ellipseLeftElbow, leftElbow, false);
            SetEllipsePosition(ellipseRightElbow, rightElbow, false);
            SetEllipsePosition(ellipseHipCenter, hipCenter, false);
            SetEllipsePosition(ellipseLeftKnee, leftKnee, false);
            SetEllipsePosition(ellipseRightKnee, rightKnee, false);
            //SetEllipsePosition(ellipseLeftFoot, leftFoot, false);
            //SetEllipsePosition(ellipseRightFoot, rightFoot, false);
            if (dualFlag)
            {
                SetEllipsePosition2(ellipseHead2, head2, false);
                SetEllipsePosition2(ellipseLeftHand2, leftHand2, isBackGestureActive2);
                SetEllipsePosition2(ellipseRightHand2, rightHand2, isForwardGestureActive2);
                SetEllipsePosition2(ellipseLeftElbow2, leftElbow2, false);
                SetEllipsePosition2(ellipseRightElbow2, rightElbow2, false);
                SetEllipsePosition2(ellipseHipCenter2, hipCenter2, false);
                SetEllipsePosition2(ellipseLeftKnee2, leftKnee2, false);
                SetEllipsePosition2(ellipseRightKnee2, rightKnee2, false);
            }

            //调用处理手势的方法
            ProcessForwardBackGesture(beforDT, head, rightHand, leftHand, leftKnee, rightKnee, hipCenter,
                                      head2, rightHand2, leftHand2, leftKnee2, rightKnee2, hipCenter2);
        }

        /// <summary>
        /// 该方法根据关节运动的跟踪数据，是用来定位的椭圆在画布上的位置 
        /// </summary>
        /// <param name="ellipse"></param>
        /// <param name="joint"></param>
        /// <param name="isHighlighted"></param>
        private void SetEllipsePosition(Ellipse ellipse, Joint joint, bool isHighlighted)
        {
            if (isHighlighted)
            {
                ellipse.Width = 60;
                ellipse.Height = 60;
                ellipse.Fill = activeBrush;
            }
            else
            {
                ellipse.Width = 20;
                ellipse.Height = 20;
                ellipse.Fill = inactiveBrush;
            }

            CoordinateMapper mapper = sensor.CoordinateMapper;

            //将三维空间坐标转化为UV平面坐标
            var point = mapper.MapSkeletonPointToColorPoint(joint.Position, sensor.ColorStream.Format);

            //调整绘制图案的位置
            Canvas.SetLeft(ellipse, point.X - ellipse.ActualWidth / 2);
            Canvas.SetTop(ellipse, point.Y - ellipse.ActualHeight / 2);
        }

        private void SetEllipsePosition2(Ellipse ellipse, Joint joint, bool isHighlighted)
        {
            if (isHighlighted)
            {
                ellipse.Width = 60;
                ellipse.Height = 60;
                ellipse.Fill = inactiveBrush;
            }
            else
            {
                ellipse.Width = 20;
                ellipse.Height = 20;
                ellipse.Fill = activeBrush;
            }

            CoordinateMapper mapper = sensor.CoordinateMapper;

            //将三维空间坐标转化为UV平面坐标
            var point = mapper.MapSkeletonPointToColorPoint(joint.Position, sensor.ColorStream.Format);

            //调整绘制图案的位置
            Canvas.SetLeft(ellipse, point.X - ellipse.ActualWidth / 2);
            Canvas.SetTop(ellipse, point.Y - ellipse.ActualHeight / 2);
        }

        /// <summary>
        /// 处理手势的方法
        /// </summary>
        /// <param name="head"></param>
        /// <param name="rightHand"></param>
        /// <param name="leftHand"></param>
        private void ProcessForwardBackGesture(DateTime beforDT, Joint head, Joint rightHand, Joint leftHand, Joint leftKnee, Joint rightKnee, Joint hipCenter,
                                               Joint head2, Joint rightHand2, Joint leftHand2, Joint leftKnee2, Joint rightKnee2, Joint hipCenter2)
        {
            //1'  左膝的y值超过髋关节的y值-0.2，触发飞踢
            if (leftKnee.Position.Y > hipCenter.Position.Y - 0.2)
            {
                if (!myFlyKick)
                {
                    myFlyKick = true;

                    DateTime afterDT = System.DateTime.Now;
                    TimeSpan ts = afterDT.Subtract(beforDT);
                    Console.WriteLine("提膝检测总共花费{0}ms.", ts.TotalMilliseconds);

                    if (!roleOrienFlag)
                    //System.Windows.Forms.SendKeys.SendWait("{Left}");
                    //2017/10/22 by zzl
                    {
                        keybd_event(68, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10); //毫秒延时
                        System.Windows.Forms.SendKeys.SendWait("{A}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{K}");
                    }
                    else
                    {
                        keybd_event(65, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10); //毫秒延时
                        System.Windows.Forms.SendKeys.SendWait("{D}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{K}");
                    }
                }
            }
            else
            {
                myFlyKick = false;
            }





            ////双手同时上举 转换角色方位
            //if ((leftHand.Position.Y > head.Position.Y - ArmRaisedThreshold) && (rightHand.Position.Y > head.Position.Y - ArmRaisedThreshold))
            //{
            //    if (!isBlackScreenActive)
            //    {
            //        keybd_event(68, 0, 2, 0);
            //        keybd_event(65, 0, 2, 0);
            //        isBlackScreenActive = true;
            //        roleOrienFlag = !roleOrienFlag;
            //    }
            //}
            //else
            //{
            //    isBlackScreenActive = false;
            //}

            //2017/11/27 by zzl
            //2' 触发重拳
            if ((head.Position.Z - rightHand.Position.Z > 0.45) || (head.Position.Z - leftHand.Position.Z > 0.45))
            {
                if (!myBoxingFlag)
                {
                    DateTime afterDT = System.DateTime.Now;
                    TimeSpan ts = afterDT.Subtract(beforDT);
                    Console.WriteLine("直拳检测总共花费{0}ms.", ts.TotalMilliseconds);

                    keybd_event(68, 0, 2, 0);
                    keybd_event(65, 0, 2, 0);
                    myBoxingFlag = true;
                    System.Windows.Forms.SendKeys.SendWait("{J}");
                }
            }
            else
            {
                myBoxingFlag = false;
            }

            //3'若右手位置的横坐标值超过设定的阈值，触发SDSDK
            if (rightHand.Position.X > head.Position.X + ArmStretchedThreshold && !isBlackScreenActive && !myFlyKick)
            {
                if (!isForwardGestureActive)
                {
                    isForwardGestureActive = true;
                    DateTime afterDT = System.DateTime.Now;

                    TimeSpan ts = afterDT.Subtract(beforDT);
                    Console.WriteLine("抬右手检测总共花费{0}ms.", ts.TotalMilliseconds);

                    //模拟键盘
                    //     System.Windows.Forms.SendKeys.SendWait("{Right}");
                    //2017/10/22 by zzl
                    if (!roleOrienFlag)
                    {
                        keybd_event(68, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{D}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{D}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{K}");
                    }
                    else
                    {
                        keybd_event(65, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{A}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{A}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{K}");
                    }
                }
            }
            else
            {
                isForwardGestureActive = false;
            }

            //4'若左手位置的横坐标超过设定的阈值，触发SDSDJ
            if (leftHand.Position.X < head.Position.X - ArmStretchedThreshold && !isBlackScreenActive && !myFlyKick)
            {
                if (!isBackGestureActive)
                {
                    isBackGestureActive = true;
                    DateTime afterDT = System.DateTime.Now;
                    TimeSpan ts = afterDT.Subtract(beforDT);
                    Console.WriteLine("抬左手检测总共花费{0}ms.", ts.TotalMilliseconds);

                    if (!roleOrienFlag)
                    //System.Windows.Forms.SendKeys.SendWait("{Left}");
                    //2017/10/22 by zzl
                    {
                        keybd_event(68, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10); //毫秒延时
                        System.Windows.Forms.SendKeys.SendWait("{D}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{D}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{J}");
                        roleOrienFlag = true;
                    }
                    else
                    {
                        keybd_event(65, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10); //毫秒延时
                        System.Windows.Forms.SendKeys.SendWait("{A}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{S}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{A}");
                        Thread.Sleep(10);
                        System.Windows.Forms.SendKeys.SendWait("{J}");
                        roleOrienFlag = false;
                    }
                }
            }
            else
            {
                isBackGestureActive = false;
            }

            //2017/11/20 by zzl
            //5' 若头部上移则触发跳跃
            if (head.Position.Y - last_head_y > 0.08)
            {
                if (!myJumpFlag)
                {
                    DateTime afterDT = System.DateTime.Now;
                    TimeSpan ts = afterDT.Subtract(beforDT);
                    Console.WriteLine("头部上移检测总共花费{0}ms.", ts.TotalMilliseconds);

                    keybd_event(68, 0, 2, 0);
                    keybd_event(65, 0, 2, 0);
                    myJumpFlag = true;
                    System.Windows.Forms.SendKeys.SendWait("{W}");
                }
            }
            else
            {
                myJumpFlag = false;
            }

            //should be sampled
            if (Framesecond == 30)
                last_head_y = head.Position.Y;

            //2017/11/14 by zzl
            //6'触发后退
            if (Math.Abs(head.Position.X - rightHand.Position.X) < 0.15 && (Math.Abs(head.Position.X - leftHand.Position.X) < 0.15) && !isForwardGestureActive && !isBackGestureActive && !myBoxingFlag)
            // if (( head.Position.Z - rightHand.Position.Z> 0.15) && ( head.Position.Z - leftHand.Position.Z > 0.15) && !myBoxingFlag)
            {
                if (!myBackFlag)
                {
                    DateTime afterDT = System.DateTime.Now;
                    TimeSpan ts = afterDT.Subtract(beforDT);
                    Console.WriteLine("双手靠近身体中心线检测总共花费{0}ms.", ts.TotalMilliseconds);

                    if (!roleOrienFlag)
                    {

                        myBackFlag = true;
                        keybd_event(68, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{A}");
                        Thread.Sleep(10);
                        keybd_event(65, 0, 1, 0);
                        Thread.Sleep(100);
                        keybd_event(65, 0, 2, 0);
                    }
                    else
                    {
                        myBackFlag = true;
                        keybd_event(65, 0, 2, 0);
                        System.Windows.Forms.SendKeys.SendWait("{D}");
                        Thread.Sleep(10);
                        keybd_event(68, 0, 1, 0);
                        Thread.Sleep(100);
                        keybd_event(68, 0, 2, 0);
                    }
                }
            }
            else
            {
                myBackFlag = false;
            }

            //2017/11/14 by zzl
            //7'触发前进
            if ((rightHand.Position.Z - head.Position.Z > 0.25) && (leftHand.Position.Z - head.Position.Z > 0.25) && !myBoxingFlag && !myBackFlag)
            {
                if (!myForwardFlag)
                {
                    DateTime afterDT = System.DateTime.Now;
                    TimeSpan ts = afterDT.Subtract(beforDT);
                    Console.WriteLine("双手侧后伸展检测总共花费{0}ms.", ts.TotalMilliseconds);

                    if (!roleOrienFlag)
                    {
                        myForwardFlag = true;
                        System.Windows.Forms.SendKeys.SendWait("{D}");
                        Thread.Sleep(10);
                        keybd_event(68, 0, 1, 0);
                        //                   Thread.Sleep(200);
                        //                   keybd_event(68, 0, 2, 0);
                    }
                    else
                    {
                        myForwardFlag = true;
                        System.Windows.Forms.SendKeys.SendWait("{A}");
                        Thread.Sleep(10);
                        keybd_event(65, 0, 1, 0);
                        //                  Thread.Sleep(200);
                        //                   keybd_event(65, 0, 2, 0);
                    }
                }
            }
            else
            {
                myForwardFlag = false;
            }

            if (dualFlag)
            {
                //角色2
                //1'  左膝的y值超过髋关节的y值，触发飞踢
                if (leftKnee2.Position.Y > hipCenter2.Position.Y - 0.2)
                {
                    if (!myFlyKick2)
                    {
                        myFlyKick2 = true;

                        if (!roleOrienFlag)
                        //System.Windows.Forms.SendKeys.SendWait("{Left}");
                        //2017/10/22 by zzl
                        {
                            keybd_event(37, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10); //毫秒延时
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            keybd_event(98, 0, 1, 0);
                            keybd_event(98, 0, 2, 0);
                        }
                        else
                        {
                            keybd_event(39, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10); //毫秒延时
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            Thread.Sleep(10);
                            keybd_event(98, 0, 1, 0);
                            keybd_event(98, 0, 2, 0);
                        }
                    }
                }
                else
                {
                    myFlyKick2 = false;
                }

                //2017/11/27 by zzl
                //2' 触发重拳
                if ((head2.Position.Z - rightHand2.Position.Z > 0.45) || (head2.Position.Z - leftHand2.Position.Z > 0.45))
                {
                    if (!myBoxingFlag2)
                    {
                        keybd_event(37, 0, 2, 0);
                        keybd_event(39, 0, 2, 0);
                        myBoxingFlag2 = true;
                        keybd_event(97, 0, 1, 0);
                        keybd_event(97, 0, 2, 0);
                    }
                }
                else
                {
                    myBoxingFlag2 = false;
                }

                //3'若右手位置的横坐标值超过设定的阈值，触发下前下前腿
                if (rightHand2.Position.X > head2.Position.X + ArmStretchedThreshold && !isBlackScreenActive && !myFlyKick2)
                {
                    if (!isForwardGestureActive2)
                    {
                        isForwardGestureActive2 = true;

                        if (!roleOrienFlag)
                        {
                            keybd_event(37, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            Thread.Sleep(10);
                            keybd_event(98, 0, 1, 0);
                            keybd_event(98, 0, 2, 0);
                        }
                        else
                        {
                            keybd_event(39, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            keybd_event(98, 0, 1, 0);
                            keybd_event(98, 0, 2, 0);
                        }
                    }
                }
                else
                {
                    isForwardGestureActive2 = false;
                }

                //4'若左手位置的横坐标超过设定的阈值，触发下后前拳
                if (leftHand2.Position.X < head2.Position.X - ArmStretchedThreshold && !isBlackScreenActive && !myFlyKick2)
                {
                    if (!isBackGestureActive2)
                    {
                        isBackGestureActive2 = true;

                        if (!roleOrienFlag)
                        //2017/10/22 by zzl
                        {
                            keybd_event(37, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10); //毫秒延时
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            Thread.Sleep(10);
                            keybd_event(97, 0, 1, 0);
                            keybd_event(97, 0, 2, 0);
                        }
                        else
                        {
                            keybd_event(39, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                            Thread.Sleep(10); //毫秒延时
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            Thread.Sleep(10);
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            keybd_event(97, 0, 1, 0);
                            keybd_event(97, 0, 2, 0);
                        }
                    }
                }
                else
                {
                    isBackGestureActive2 = false;
                }

                //2017/11/20 by zzl
                //5' 若头部上移则触发跳跃
                if (head2.Position.Y - last_head_y2 > 0.08)
                {
                    if (!myJumpFlag2)
                    {
                        keybd_event(37, 0, 2, 0);
                        keybd_event(39, 0, 2, 0);
                        myJumpFlag2 = true;
                        System.Windows.Forms.SendKeys.SendWait("{UP}");
                    }
                }
                else
                {
                    myJumpFlag2 = false;
                }

                //should be sampled
                if (Framesecond == 30)
                    last_head_y2 = head2.Position.Y;

                //2017/11/14 by zzl
                //6'触发后退
                if (Math.Abs(head2.Position.X - rightHand2.Position.X) < 0.15 && (Math.Abs(head2.Position.X - leftHand2.Position.X) < 0.15) && !isForwardGestureActive2 && !isBackGestureActive2 && !myBoxingFlag2)
                // if (( head.Position.Z - rightHand.Position.Z> 0.15) && ( head.Position.Z - leftHand.Position.Z > 0.15) && !myBoxingFlag)
                {
                    if (!myBackFlag2)
                    {
                        if (!roleOrienFlag)
                        {
                            myBackFlag2 = true;
                            keybd_event(37, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            keybd_event(39, 0, 1, 0);
                            Thread.Sleep(100);
                            keybd_event(39, 0, 2, 0);
                        }
                        else
                        {
                            myBackFlag2 = true;
                            keybd_event(39, 0, 2, 0);
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            keybd_event(37, 0, 1, 0);
                            Thread.Sleep(100);
                            keybd_event(37, 0, 2, 0);
                        }
                    }
                }
                else
                {
                    myBackFlag2 = false;
                }

                //2017/11/14 by zzl
                //7'触发前进
                if ((rightHand2.Position.Z - head2.Position.Z > 0.25) && (leftHand2.Position.Z - head2.Position.Z > 0.25) && !myBoxingFlag2 && !myBackFlag2)
                {
                    if (!myForwardFlag2)
                    {
                        if (!roleOrienFlag)
                        {
                            myForwardFlag2 = true;
                            System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                            Thread.Sleep(10);
                            keybd_event(37, 0, 1, 0);
                            //                   Thread.Sleep(200);
                            //                   keybd_event(68, 0, 2, 0);
                        }
                        else
                        {
                            myForwardFlag2 = true;
                            System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(10);
                            keybd_event(39, 0, 1, 0);
                            //                  Thread.Sleep(200);
                            //                   keybd_event(65, 0, 2, 0);
                        }
                    }
                }
                else
                {
                    myForwardFlag2 = false;
                }

            }


            //判断双手靠近头部
            if (Math.Abs(head.Position.Y - rightHand.Position.Y) < DistanceThreshold && (Math.Abs(head.Position.Y - leftHand.Position.Y) < DistanceThreshold && !isForwardGestureActive && !isBackGestureActive))
            {
                if (!isPresent)
                {
                    isPresent = true;
                    System.Windows.Forms.SendKeys.SendWait("{F5}");
                }
            }
            else
            {
                isPresent = false;
            }
        }

        /// <summary>
        /// 切换圆形图案的方法
        /// </summary>
        void ToggleCircles()
        {
            if (isCirclesVisible)
                HideCircles();
            else
                ShowCircles();
        }

        /// <summary>
        /// 隐藏圆图案的方法
        /// </summary>
        void HideCircles()
        {
            isCirclesVisible = false;
            // ellipseHead.Visibility = System.Windows.Visibility.Collapsed;
            //ellipseLeftHand.Visibility = System.Windows.Visibility.Collapsed;
            // ellipseRightHand.Visibility = System.Windows.Visibility.Collapsed;
            //2017/11/29 by zzl
            ellipseLeftHand2.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightHand2.Visibility = System.Windows.Visibility.Collapsed;
            ellipseHead2.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftElbow2.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightElbow2.Visibility = System.Windows.Visibility.Collapsed;
            ellipseHipCenter2.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftKnee2.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightKnee2.Visibility = System.Windows.Visibility.Collapsed;


        }

        /// <summary>
        /// 显示圆形图案的方法
        /// </summary>
        void ShowCircles()
        {
            isCirclesVisible = true;
            //  ellipseHead.Visibility = System.Windows.Visibility.Visible;
            //  ellipseLeftHand.Visibility = System.Windows.Visibility.Visible;
            //  ellipseRightHand.Visibility = System.Windows.Visibility.Visible;
            //2017/11/29 by zzl
            ellipseLeftHand2.Visibility = System.Windows.Visibility.Visible;
            ellipseRightHand2.Visibility = System.Windows.Visibility.Visible;
            ellipseHead2.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftElbow2.Visibility = System.Windows.Visibility.Visible;
            ellipseRightElbow2.Visibility = System.Windows.Visibility.Visible;
            ellipseHipCenter2.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftKnee2.Visibility = System.Windows.Visibility.Visible;
            ellipseRightKnee2.Visibility = System.Windows.Visibility.Visible;
        }

        /// <summary>
        /// 显示窗体的方法
        /// </summary>
        private void ShowWindow()
        {
            this.Topmost = true;
            this.WindowState = System.Windows.WindowState.Maximized;
        }

        /// <summary>
        /// 隐藏窗体的方法
        /// </summary>
        private void HideWindow()
        {
            this.Topmost = false;
            this.WindowState = System.Windows.WindowState.Minimized;
        }

        #region Speech Recognition Methods

        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        private void InitializeSpeechRecognition()
        {
            RecognizerInfo ri = GetKinectRecognizer();
            if (ri == null)
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
 Ensure you have the Microsoft Speech SDK installed.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                speechRecognizer = new SpeechRecognitionEngine(ri.Id);
            }
            catch
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
 Ensure you have the Microsoft Speech SDK installed and configured.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            var phrases = new Choices();
            phrases.Add("computer show window");
            phrases.Add("computer hide window");
            phrases.Add("computer show circles");
            phrases.Add("computer hide circles");
            phrases.Add("single player");
            phrases.Add("two players");



            var gb = new GrammarBuilder();
            //Specify the culture to match the recognizer in case we are running in a different culture.                                 
            gb.Culture = ri.Culture;
            gb.Append(phrases);

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

            speechRecognizer.LoadGrammar(g);
            speechRecognizer.SpeechRecognized += SreSpeechRecognized;
            speechRecognizer.SpeechHypothesized += SreSpeechHypothesized;
            speechRecognizer.SpeechRecognitionRejected += SreSpeechRecognitionRejected;

            this.readyTimer = new DispatcherTimer();
            this.readyTimer.Tick += this.ReadyTimerTick;
            this.readyTimer.Interval = new TimeSpan(0, 0, 4);
            this.readyTimer.Start();

        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.StartSpeechRecognition();
            this.readyTimer.Stop();
            this.readyTimer.Tick -= ReadyTimerTick;
            this.readyTimer = null;
        }

        private void StartSpeechRecognition()
        {
            if (sensor == null || speechRecognizer == null)
                return;

            var audioSource = this.sensor.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();

            speechRecognizer.SetInputToAudioStream(
                    kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);

        }

        void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            Trace.WriteLine("\nSpeech Rejected, confidence: " + e.Result.Confidence);
        }

        void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            Trace.Write("\rSpeech Hypothesized: \t{0}", e.Result.Text);
        }

        void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //This first release of the Kinect language pack doesn't have a reliable confidence model, so 
            //we don't use e.Result.Confidence here.
            if (e.Result.Confidence < 0.70)
            {
                Trace.WriteLine("\nSpeech Rejected filtered, confidence: " + e.Result.Confidence);
                return;
            }

            Trace.WriteLine("\nSpeech Recognized, confidence: " + e.Result.Confidence + ": \t{0}", e.Result.Text);

            if (e.Result.Text == "computer show window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                    {
                        ShowWindow();
                    });
            }
            else if (e.Result.Text == "computer hide window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    HideWindow();
                });
            }
            else if (e.Result.Text == "computer hide circles")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.HideCircles();
                });
            }
            else if (e.Result.Text == "computer show circles")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.ShowCircles();
                });
            }
            else if (e.Result.Text == "single player")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    dualFlag = false;
                    this.HideCircles();
                });
            }
            else if (e.Result.Text == "two players")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    dualFlag = true;
                    this.ShowCircles();
                });
            }
        }
    }
}
#endregion