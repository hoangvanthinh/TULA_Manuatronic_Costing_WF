using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using EasyModbus;

using System.Drawing.Imaging;

///lres

namespace How_to_create_HMI_Control_Real_Time
{
    public partial class FormMain : Form
    {
        ModbusClient MB;
        const int ADD_X = 1;
        const int ADD_Y = 2;
        const int ADD_Z = 3;

        const int ADD_SPRAY = 4;
        const int ADD_HOME = 6;
        const int ADD_AM = 7;


        const int ADD_REQ_X = 11;
        const int ADD_REQ_Y = 12;
        const int ADD_REQ_Z = 13;


        //======================== REQ UPLOAD DATA =============================
        const int ADD_REQ_DOWLOAD = 20;
        const int ADD_SOHANHTIRNH_DOWLOAD = 21;

        const int ADD_hanhtrinh_so = 22;
        const int ADD_TYPE_HANHTRINH = 23;
        const int ADD_X1 = 24;
        const int ADD_Y1 = 25;
        const int ADD_Z1 = 26;

        const int ADD_X2 = 27;
        const int ADD_Y2 = 28;
        const int ADD_Z2 = 29;

        const int ADD_STEPS = 30;
        const int ADD_SPEED = 31;
        const int ADD_LIFT_UP = 32;
        const int ADD_EN_DISPENSER = 33;

        const int ADD_Z_MOVE = 34;
        const int ADD_XWP = 35;
        const int ADD_YWP = 36;
        const int ADD_ZWP = 37;

        const int ADD_XDSP = 38;
        const int ADD_YDSP = 39;
        const int ADD_ZDSP = 40;

        const int ADD_TIME_DROP = 41;

        const int ADD_deX = 42;
        const int ADD_deY = 43;
        const int ADD_numcopy_array = 44;
        const int ADD_ID_start_copy = 45;

        const int ADD_REPORT_FILE_LOAD_OK = 46;
        //=======================================================================

        const int LOW =1;
        const int MID =2;
        const int HIGH = 3;

        const int SCALE_MAP = 68;  //30000/450
        //----------------------------------- DU LIEU TOA DO --------------------
        int ID_HANHTRINH = 0;
        const int POINT = 1;
        const int LINE = 2;
        const int Rectangular_Array = 3;

        int TYPE_HANHTRINH = POINT;

        const int CHOICE_TOADO_1 = 1;
        const int CHOICE_TOADO_2 = 2;
        int choice_toado = 0;

        int Add_START = 100;
        int SIZE_DATA = 16;

        ushort SO_HANHTRINH = 0;
        //----------------------------------- actua ------------------------------
        int _runX = 0, _runY = 0, _runZ = 0;
        int reqX = 0, reqY = 0, reqZ = 0;
        int _AM = 0;
        const int _Semi_auto = 0,_Auto = 1, _Manual = 2;
        const int _X = 1, _Y = 2, _Z = 3;

        const int _HOME = 6, _Add_AutoManual = 7, _Add_UPS = 8, _Add_COOD = 9, _Add_SPEED = 10;

        int Zmove = 0, Zlimit = 0, PW_X = 0, PW_Y = 0, PW_Z = 0;
        const int Zmovewaiting = 500;

        int Dsp_X = 0, Dsp_Y = 0, Dsp_Z = 0;// tọa độ phun xả
        int Time_Drop = 0, time_drop_temp = 0;
        int OF1x = 0, OF1y = 0, OFnx = 0, OFny = 0;
        int Numbercopy = 0, Id_start_copy = 0;
        int _SPEED = 0;
        int i_SPEED = 1;
        const int SPEED_LOW = 0, SPEED_HIGH = 1;
        ushort _ups = 0, _coodinate = 0;
        ushort Start = 0;
        int Spray_ON = 1, Spray_OFF = 0;
        int _dispenser = 0;
        int _Status_SPRAY = 2;
        ushort AUTO_RUN = 0, _READY_RUN = 0, _Req_ORG = 0;
        int REQ_DISSPRAY = 0, _step_disspray = 0;

        //--------------------- di chuyen tu dong--------------------------
        ushort _PAUSE = 0;
        ushort _coodinateX = 0, _coodinateY = 0, _coodinateZ = 0;
        int HANHTRINH_COMPELETE = 0;
        int hanhtrinh = 0, hanhtrinh_so = 0;
        int _step = 0;
        int time_spray = 0, Timer_spray_ON = 0;
        int[] Type_spray = new int[100];  // 1- POINT 2- LINE

        int[] X1_spray = new int[100]; // X
        int[] Y1_spray = new int[100]; // Y
        int[] Z1_spray = new int[100]; // Z
        int[] X2_spray = new int[100]; // X
        int[] Y2_spray = new int[100]; // Y
        int[] Z2_spray = new int[100]; // Z
        int[] STEP_spray = new int[100]; // Z
        int[] SPEED_spray = new int[100];
        int[] LIFT_UP = new int[100];
        int[] DISPEN = new int[100];

        int step_num_array = 1;
        int _stepY = 0;
        int Dir_step = 0;
        int _tangbacXY = 1;

        int deX = 0, deY = 0,_cell_number = 1;
        int _dStepX=0;   
        int _dStepY=0;
        struct D_Calib
        {
            public int dX;
            public int dY;
            public int dZ;

        };
        D_Calib Co_Calib = new D_Calib();
        // ----------------- SAVE DATA -------------------------------
        byte[] data = new byte[1024];

        //private Register _Register = null;
        DataTable table = new DataTable();
        int row_selected = 0, cell_selected = 0;

        List<string> _items = new List<string>();

        public FormMain()
        {
            InitializeComponent();

            //Line line = Line.FormMain(new Point(0, 0), new Point(3, 4));
        }

        Bitmap TULA = new Bitmap(450, 450);
        Pen PenRed = new Pen(System.Drawing.Color.Red, 2);
        
        private void FormMain_Load(object sender, EventArgs e)
        {
            try
            {
                
                table.Columns.Add("ID", typeof(uint));// data type int
                table.Columns.Add("TYPE", typeof(string));// datatype string

                table.Columns.Add("X1", typeof(uint));// datatype string
                table.Columns.Add("Y1", typeof(uint));// datatype string
                table.Columns.Add("Z1", typeof(uint));// data type int

                table.Columns.Add("X2", typeof(uint));// datatype string
                table.Columns.Add("Y2", typeof(uint));// datatype string
                table.Columns.Add("Z2", typeof(uint));// data type int        

                table.Columns.Add("Steps", typeof(uint));// data type int 
                table.Columns.Add("Speed", typeof(uint));// data type int 
                table.Columns.Add("Lift up", typeof(bool));
                table.Columns.Add("Dispenser", typeof(bool));
                DT.DataSource = table;
                DT.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                
                MB = new ModbusClient("COM42");
                MB.UnitIdentifier = 1;
                MB.Baudrate = 115200;
                MB.Parity = System.IO.Ports.Parity.None;
                MB.StopBits = System.IO.Ports.StopBits.One;
                MB.ConnectionTimeout = 1000;
                MB.Connect();
                //============= doc trang thai dieu khien =========================//
                _AM = MB.ReadHoldingRegisters(ADD_AM, 1)[0];
                //if (_AM == 0)
                //    Automanual.Text = "Semi_auto";
                //if (_AM == _Manual)
                //    Automanual.Text = "Manual";
                //if(_AM == _Auto)
                //    Automanual.Text = "Auto";
                CHECK_AM();
                Set_speed_manual.Value = 2;
                Speed_manual.Text = Set_speed_manual.Value.ToString();
                //============ DOC GIA TRI KHOI TAO TOC DO ========================//
                _SPEED = MB.ReadHoldingRegisters(10, 1)[0];
                i_SPEED = _SPEED;
                if (_SPEED == LOW)
                    Select_Speed.Text = "LOW SPEED";
                if (_SPEED == MID)
                    Select_Speed.Text = "MID SPEED";
                if (_SPEED == HIGH)
                    Select_Speed.Text = "HIGH SPEED";

                //================================================================//
                _Status_SPRAY = MB.ReadHoldingRegisters(ADD_SPRAY, 1)[0];
                tocdo_truc.Value = 2;
                MB.WriteSingleRegister(14, tocdo_truc.Value);

                if (_Status_SPRAY == Spray_ON)
                    PUMP.BackColor = Color.Green;
                if (_Status_SPRAY == Spray_OFF)
                    PUMP.BackColor = Color.Red;
                //===================== KHOI TAO SET TOA DO =======================//
                //ID_HANHTRINH = DT.RowCount;
                ID_PL.Text = ID_HANHTRINH.ToString();

                TYPE_HANHTRINH = LINE;
                Type_PL.Text = "Line";

                NB_steps.Hide();
                Number_step.Hide();

                X1.Enabled = false;
                Y1.Enabled = false;
                Z1.Enabled = false;
                X2.Enabled = false;
                Y2.Enabled = false;
                Z2.Enabled = false;
                //=================== BAT TIMER ===================================//
                timer1.Enabled = true;

               // DT.CancelEdit();
                linkLabel2.Text = "Click here to get more info.";
                linkLabel2.Links.Add(6,4,"www.tula.vn");
                


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            check_Sequential.Checked = true;
            
            //================== RECALL PROJECT ==============================
            string Projectname = "";
            using (StreamReader SR = new StreamReader("Name_project.txt"))
            {
                while ((Projectname = SR.ReadLine()) != null)
                    List_file.Items.Add(Projectname);
            }

        }


        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            MessageBox.Show("KeyDown");
        }


        private void HOME_Click(object sender, EventArgs e)
        {
            Manual.Enabled = false;

            i_SPEED = MID;
            Select_Speed.Text = "MID SPEED";

            _Status_SPRAY = Spray_OFF;
            PUMP.BackColor = Color.Red;

            Automanual.Text = "Manual";
            try
            {
                Req_GOHOME();
            }
            catch //(Exception ex)
            {

                MB.Connect();
            }
     
            
        }

        private void Req_GOHOME()
        {
            //throw new NotImplementedException();
            try
            {
                MB.WriteSingleRegister(_HOME, 1); // Max= 65535
            }
            catch
            {

            }
        }

        private void button5_MouseDown(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 1;
                _coodinate = _X;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }
            }
        }

        private void button5_MouseUp(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 0;
                _coodinate = _X;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button4_MouseDown(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 2;
                _coodinate = _X;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button4_MouseUp(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 0;
                _coodinate = _X;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }


        private void Y_UP_MouseDown(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 1;
                _coodinate = _Y;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void Y_UP_MouseUp(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 0;
                _coodinate = _Y;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button3_MouseDown(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 2;
                _coodinate = _Y;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button3_MouseUp(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 0;
                _coodinate = _Y;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button7_MouseDown(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 1;
                _coodinate = _Z;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button7_MouseUp(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 0;
                _coodinate = _Z;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button6_MouseDown(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 2;
                _coodinate = _Z;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }

            }
        }

        private void button6_MouseUp(object sender, MouseEventArgs e)
        {
            if (_AM == _Manual)
            {
                _ups = 0;
                _coodinate = _Z;
                try
                {
                    MB.WriteSingleRegister(_Add_COOD, _coodinate);
                    MB.WriteSingleRegister(_Add_UPS, _ups);
                }
                catch
                {

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System_config systemCF = new System_config();
            systemCF.Show();
        }
        private void Draw_hanhtrinh(int x1, int y1, int x2, int y2)
        {

            Bitmap tmp = new Bitmap(450, 450);
            Graphics g1 = Graphics.FromImage(tmp);
            g1.DrawLine(PenRed, x1 / SCALE_MAP, y1 / SCALE_MAP, x2 / SCALE_MAP, y2 / SCALE_MAP);

            TULA = MergedBitmaps(TULA, tmp);
            MAP.BackgroundImage = TULA;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            if (Timer_spray_ON == 1)
                time_spray++;
           

            if (!MB.Connected)
            {
                if (MB.Available(50))
                {
                    //timer1.Enabled = false;
                    //MB.Disconnect();
                    //MB.Connect();
                    timer1.Enabled = true;
                }
            }
            else
            {
                //Thread tula = new Thread(Chay_timer);
                //tula.IsBackground = true;
                //tula.Start();
                Chay_timer();
            }
        }
        void Chay_timer()
        {
            
            try
            {
                
                Manual.Enabled = true;

       
                _coodinateX = (ushort)MB.ReadHoldingRegisters(ADD_X, 1)[0];
                _coodinateY = (ushort)MB.ReadHoldingRegisters(ADD_Y, 1)[0];
                _coodinateZ = (ushort)MB.ReadHoldingRegisters(ADD_Z, 1)[0];
                
                reqX = MB.ReadHoldingRegisters(ADD_REQ_X, 1)[0];
                reqY = MB.ReadHoldingRegisters(ADD_REQ_Y, 1)[0];
                reqZ = MB.ReadHoldingRegisters(ADD_REQ_Z, 1)[0];

                if (choice_toado == 0)
                {
                    Toado_X.Text = _coodinateX.ToString();
                    Toado_Y.Text = _coodinateY.ToString();
                    Toado_Z.Text = _coodinateZ.ToString();
                }
                else if (choice_toado == CHOICE_TOADO_1)
                {
                    X1.Text = _coodinateX.ToString();
                    Y1.Text = _coodinateY.ToString();
                    Z1.Text = _coodinateZ.ToString();

                }
                else
                {
                    X2.Text = _coodinateX.ToString();
                    Y2.Text = _coodinateY.ToString();
                    Z2.Text = _coodinateZ.ToString();
                }

                MB.WriteSingleRegister(5, Zlimit);

                Tocdo_dichuyen.Text = MB.ReadHoldingRegisters(10, 1)[0].ToString();
                //_SPR.Text = MB.ReadHoldingRegisters(ADD_SPRAY, 1)[0].ToString();

                if (AUTO_RUN == 1 && _READY_RUN == 0 && _Req_ORG == 1)
                {
                    _Req_ORG = 0;
                    MB.WriteSingleRegister(_Add_SPEED, SPEED_HIGH);
                    try
                    {
                        Req_GOHOME();
                    }
                    catch
                    {

                        MB.Connect();
                    }

                }
                if (AUTO_RUN == 1 && _coodinateX == 0 && _coodinateY == 0 && _coodinateZ == 0)
                {
                    //=========>> BAT DAU CHAY AUTO  <<===================================//
                    _READY_RUN = 1;
                    MB.WriteSingleRegister(ADD_AM, _Auto);

                    Automanual.Text = "Auto";
                    //================== khoi tao Hanh trinh =============================
                    _step = 1;
                    hanhtrinh_so = 1;
                    REQ_DISSPRAY = 1;
                    _step_disspray = 1;

                }
                if (AUTO_RUN == 1 && _READY_RUN == 1 && _PAUSE == 0)
                {
                    //===============================KIEM TRA TRANG THAI CAC TRUC 
                    if (reqX == _coodinateX && reqY == _coodinateY)
                    {
                        _runX = _runY = 0;
                    }
                    if (reqZ == _coodinateZ)
                        _runZ = 0;

                    //========================= DI CHUYEN CHAM KEO   ==========>>>>
                    if (HANHTRINH_COMPELETE == 0)
                    {

                        // CHAY HANH TRINH THEO TOA DO
                        CHAY_HANHTRINH();
                    }

                }
                if (REQ_DISSPRAY == 1)
                    time_drop_temp++;
                timer1.Enabled = true;
              
            }
            catch
            {
                //MB.Disconnect();
                //MB.Connect();
                timer1.Enabled = true;
            }
      
        }
        void Sum_dxdy()
        {
            _dStepX = deX * (_cell_number - 1) / (Numbercopy - 1);
            _dStepY = deY * (_cell_number - 1) / (Numbercopy - 1);
            //textBox3.Text = _cell_number.ToString();
            //textBox4.Text = _dStepX.ToString();
            //textBox5.Text = _dStepY.ToString();
        }
            
        private void CHAY_HANHTRINH()
        {
            #region XA BAN DAU
            if (REQ_DISSPRAY == 1)
            {
                if(_step_disspray == 1)
                {
                    time_drop_temp = 0;  // reset bien thoi gian
                    MoveXY(Dsp_X, Dsp_Y);
                    _runX = _runY = 1;
                    _step_disspray++;

                }
                else if(_step_disspray == 2 && _runX == 0 && _runY == 0)
                {
                    MoveZ(Dsp_Z);
                    _runZ = 1;
                    _step_disspray++;
                }
                else if(_step_disspray == 3 && _runZ == 0)
                {
                    SPRAY_ENABLE();
                    if(time_drop_temp >= Time_Drop*10) 
                        _step_disspray++;
                }
                else if(_step_disspray == 4)
                {
                    SPRAY_DISABLE();
                    _step_disspray++;
                }
                else if(_step_disspray == 5)
                {
                    MoveZ(Zmove);
                    _runZ = 1;
                    _step_disspray++;
                }
                else if (_step_disspray == 6 && _runZ == 0)
                {
                    _step_disspray = 1;
                    REQ_DISSPRAY = 0;
                }


            }
            #endregion
            //_dStepX = deX * (_cell_number - 1) / (Numbercopy - 1);
            //_dStepX = deY * (_cell_number - 1) / (Numbercopy - 1);
            
            if (hanhtrinh_so < Id_start_copy && _cell_number > 1)
                hanhtrinh_so = Id_start_copy;
            #region DIEM
            if (Type_spray[hanhtrinh_so] == POINT && REQ_DISSPRAY == 0) //============== ĐIỂM ================================================
            {
                Sum_dxdy();
                if (_step == 1)  
                {
                    //MoveXY(X1_spray[hanhtrinh_so], Y1_spray[hanhtrinh_so]);
                    MoveXY(X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + +_dStepY);
                    _runX = _runY = 1;
                    _step++;
                }
                else if (_step == 2 && _runX == 0 && _runY == 0) // XUONG NOZZLE
                {
                    MoveZ(Z1_spray[hanhtrinh_so]);
                    _runZ = 1;
                    _step++;
          
                }
                else if (_step == 3 && _runZ == 0)// PHUN
                {
                    SPRAY_ENABLE();
                    Timer_spray_ON = 1;  // BAT TIMER SPRAY
                    time_spray = 0;
                    _runZ = 1;
                    _step++;

                }
                else if(_step == 4 && time_spray >= 2)// 50mS*10
                {
                    SPRAY_DISABLE();
                    Timer_spray_ON = 0;
                    _step++;
                }
                else if(_step == 5)
                {
                    if (LIFT_UP[hanhtrinh_so] == 0)
                    {
                        _runZ = 0;
                    }
                    else
                    {
                        MoveZ(Zmove);
                        _runZ = 1;
                    }

                    _step++;
                }
                else if(_step == 6 && _runZ == 0)
                {
                    //if (hanhtrinh_so < hanhtrinh)
                    //{
                    //    hanhtrinh_so++;
                    //    _step = 1;
                    //}
                    //else
                    //{
                    //    MoveZ(Zmovewaiting);
                    //    _step++;

                    //}
                    if (hanhtrinh_so < hanhtrinh)
                    {
                        hanhtrinh_so++;
                        _step = 1;
                    }
                    else
                    {
                        if (_cell_number >= Numbercopy)
                        {
                            MoveZ(Zmovewaiting);
                            _runZ = 1;
                            _step++;
                        }
                        else
                        {
                            _cell_number++;
                            hanhtrinh_so = 1;
                            _step = 1;
                        }
                    } 
                }
                else if(_step == 7 && _runZ == 0)
                {
                    MoveXY(PW_X, PW_Y);
                    _step++;

                }
                else if(_step == 8 && _runX == 0 && _runY ==0)
                {
                    MoveZ(PW_Z);
                    hanhtrinh_so = 0;
                    _step = 1;
                }
            }
            #endregion 
            #region DUONG
            if (Type_spray[hanhtrinh_so] == LINE && REQ_DISSPRAY == 0) //=============== ĐƯỜNG ========================================================
            {
                Sum_dxdy();

                if (_step == 1)  
                {
                    if (hanhtrinh_so > 1)
                        DT.Rows[hanhtrinh_so - 2].Selected = false;
                    DT.Rows[hanhtrinh_so - 1].Selected = true;

                    //MoveXY(X1_spray[hanhtrinh_so] + deX * (_cell_number - 1) / (Numbercopy - 1), Y1_spray[hanhtrinh_so] + deY * (_cell_number - 1) / (Numbercopy - 1));
                    MoveXY(X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY);
                    
                    _runX = _runY = 1;
                    _step++;
                }
                else if (_step == 2 && _runX == 0 && _runY == 0)
                {
                    MoveZ(Z1_spray[hanhtrinh_so]);
                    _runZ = 1;
                    _step++;
                }
                else if(_step == 3 && _runZ == 0)
                {
                    SPRAY_ENABLE();
                    _step++;
                }
                else if(_step == 4)
                {
                   //MoveXY(X2_spray[hanhtrinh_so] + deX * (_cell_number - 1) / (Numbercopy - 1), Y2_spray[hanhtrinh_so] + deY * (_cell_number - 1) / (Numbercopy - 1));
                    MoveXY(X2_spray[hanhtrinh_so] + _dStepX, Y2_spray[hanhtrinh_so] + _dStepY);

                   //Draw_hanhtrinh(X1_spray[hanhtrinh_so] + deX * (_cell_number - 1) / (Numbercopy - 1), Y1_spray[hanhtrinh_so] + deY * (_cell_number - 1) / (Numbercopy - 1), X2_spray[hanhtrinh_so] + deX * (_cell_number - 1) / (Numbercopy - 1), Y2_spray[hanhtrinh_so] + deY * (_cell_number - 1) / (Numbercopy - 1));
                    Draw_hanhtrinh(X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY, X2_spray[hanhtrinh_so] + _dStepX, Y2_spray[hanhtrinh_so] + _dStepY);

                    _runX = _runY = 1;
                    _step++;
       
                }
                else if(_step == 5 && _runX == 0 && _runY == 0)
                {

                    try
                    {
                        SPRAY_DISABLE();
                        _step++;
                    }
                    catch
                    {
                        MessageBox.Show("Don't SPR");
                        _step++;
                    }
                                                     
                }
                else if(_step == 6)//=================== LEN NOZZLE ======
                {
                   
                    //if ((hanhtrinh_so < hanhtrinh) &&(X2_spray[hanhtrinh_so] == X1_spray[hanhtrinh_so + 1]) && (Y2_spray[hanhtrinh_so] == Y1_spray[hanhtrinh_so + 1]))
                    if (LIFT_UP[hanhtrinh_so] == 0)
                    {
                        _runZ = 0;
                    }
                    else
                    {
                        MoveZ(Zmove);
                        _runZ = 1;  
                    }
                
                    _step++;
                    
                }
                else if(_step == 7 && _runZ == 0)
                {

                    if (hanhtrinh_so < hanhtrinh)
                    {
                        hanhtrinh_so++;
                        _step = 1;
                    }
                    else
                    {
                        if (_cell_number >= Numbercopy)
                        {
                            MoveZ(Zmovewaiting);
                            _runZ = 1;
                            _step++;
                        }
                        else
                        {
                            _cell_number++;
                            hanhtrinh_so = 1;
                            _step = 1;
                           // MessageBox.Show(_cell_number.ToString() + "celln");
                        }
                    } 
                    
                }

                else if (_step == 8 && _runZ == 0)//=========================== ket thuc toan bo hanh trinh ==================================
                {
                    MoveXY(PW_X, PW_Y);
                    _runX = _runY = 1;
                    _step++;
                }
                else if(_step == 9 && _runX == 0 && _runY == 0)
                {
                    MoveZ(PW_Z);
                    _step = 1;
                    hanhtrinh_so = 0;
                }
            }
            #endregion 
            #region Array
            //======================================= Rectangular_Array ================================
            if (Type_spray[hanhtrinh_so] == Rectangular_Array && REQ_DISSPRAY == 0)
            {
                Sum_dxdy();
                if (Y1_spray[hanhtrinh_so] < Y2_spray[hanhtrinh_so])
                    Dir_step = 1;
                else
                    Dir_step = -1;
                if (_step == 1)
                {
                    if (hanhtrinh_so > 1)
                        DT.Rows[hanhtrinh_so - 2].Selected = false;
                    DT.Rows[hanhtrinh_so - 1].Selected = true;

                    MoveXY(X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY);
                     
                    _runX = _runY = 1;
                    _step++;
                }
                else if (_step == 2 && _runX == 0 && _runY == 0)
                {
                    MoveZ(Z1_spray[hanhtrinh_so]);
                    _runZ = 1;
                    _step++;
                }
                else if (_step == 3 && _runZ == 0)
                {
                    SPRAY_ENABLE();
                    _step++;
                }

                else if(_step == 4)
                {
                    MoveXY(X2_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY);
                    Draw_hanhtrinh(_coodinateX + _dStepX, _coodinateY + _dStepY, X2_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY);
                    _runX = _runY = 1;
                    _step++;
                }
                //-================== lap lai  X TO Y  ==============================
  
                else if(_step == 5)// && _runX == 0 && _runY == 0)
                {
                        if (_runX == 0 && _runY == 0 && _tangbacXY == 1)
                        {
                            if (step_num_array % 2 == 1)
                            {
                                MoveXY(X2_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY + Dir_step * step_num_array * _stepY);  // hoac X1
                                Draw_hanhtrinh(_coodinateX + _dStepX, _coodinateY + _dStepY, X2_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY+ Dir_step * step_num_array * _stepY);
                            }
                            else
                            {
                                MoveXY(X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY + Dir_step * step_num_array * _stepY);
                                Draw_hanhtrinh(_coodinateX + _dStepX, _coodinateY + _dStepY, X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY + Dir_step * step_num_array * _stepY);
                            }
                            _runX = _runY = 1;                          
                            _tangbacXY = 2;
                        }
                        if(_runX == 0 && _runY == 0 && _tangbacXY == 2)
                        {
                            if (step_num_array % 2 == 1)
                            {
                                MoveXY(X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY + Dir_step * step_num_array * _stepY);
                                Draw_hanhtrinh(_coodinateX + _dStepX, _coodinateY + _dStepY, X1_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY + Dir_step * step_num_array * _stepY);
                            }
                            else
                            {
                                MoveXY(X2_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY + Dir_step * step_num_array * _stepY);
                                Draw_hanhtrinh(_coodinateX + _dStepX, _coodinateY + _dStepY, X2_spray[hanhtrinh_so] + _dStepX, Y1_spray[hanhtrinh_so] + _dStepY + Dir_step * step_num_array * _stepY);
                            }
                            _runX = _runY = 1;
                            _tangbacXY = 1;
                            step_num_array++;
                        }
                        if (step_num_array > STEP_spray[hanhtrinh_so])
                        {
                            step_num_array = 1;
                            _step++;
                            
                        }                  
                }
                else if (_step == 6 && _runX == 0 && _runY == 0)
                {
                    try
                    {
                        
                        SPRAY_DISABLE();
                        _step++;
   
                    }
                    catch
                    {
                  
                        MessageBox.Show("Don't SPR");
                        _step++;
                    }
                }
                else if (_step == 7 )// ======================LEN NOZZLE ===========================
                {

                    if (LIFT_UP[hanhtrinh_so] == 0)
                    {
                        _runZ = 0;
                    }
                    else
                    {
                        MoveZ(Zmove);
                        _runZ = 1;
                    }
                    _step++;
                }
              
                else if (_step == 8 && _runZ == 0)
                {
                    if (hanhtrinh_so < hanhtrinh)
                    {
                        hanhtrinh_so++;
                        _step = 1;
                    }
                    else
                    {
                        if (_cell_number >= Numbercopy)
                        {
                            MoveZ(Zmovewaiting);
                            _runZ = 1;
                            _step++;
                        }
                        else
                        {
                            _cell_number++;
                            hanhtrinh_so = 1;
                            _step = 1;
                        }
                    }                    
                }

                else if (_step == 9 && _runZ == 0) //===================== ket thuc toan bo hanh trinh ===========================
                {
                    MoveXY(PW_X, PW_Y);
                    _runX = _runY = 1;
                    _step++;
                }
                else if(_step == 10 && _runX ==0 && _runY == 0)
                {
                    MoveZ(PW_Z);
                    hanhtrinh_so = 0;
                    _step = 1;
                }

            }
            #endregion
        }

        private void SPRAY_ENABLE()
        {
            if(_dispenser == 1)
                MB.WriteSingleRegister(ADD_SPRAY, Spray_ON);
        }
        private void SPRAY_DISABLE()
        {
            MB.WriteSingleRegister(ADD_SPRAY, Spray_OFF);
        }
      
        private void MoveZ(int cood_z)
        {
            //throw new NotImplementedException();
            try
            {
                MB.WriteSingleRegister(ADD_REQ_Z, cood_z);
            }
            catch
            {

            }
        }

        private void MoveXY(int cood_x, int cood_y)
        {

            //throw new NotImplementedException();
            try
            {
                MB.WriteSingleRegister(ADD_REQ_X, cood_x);
                MB.WriteSingleRegister(ADD_REQ_Y, cood_y);
            }
            catch
            {

            }
        }

     
      void IC_but_DoubleClick(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Label lab = (Label)sender;
            if (lab.BackColor == Color.Green)
                lab.BackColor = Color.White;
            else
                lab.BackColor = Color.Green;

        }
        void ROM_but_DoubleClick(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Label lab = (Label)sender;
            if (lab.BackColor == Color.Green)
                lab.BackColor = Color.White;
            else
                lab.BackColor = Color.Green;
        }

        void NG_but_DoubleClick(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Label lab = (Label)sender;
            if (lab.BackColor == Color.White)
                lab.BackColor = Color.Red;
            else
            {
                lab.BackColor = Color.White;
            }
        }


        void Pass_but_DoubleClick(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Label lab = (Label)sender;
            if (lab.BackColor == Color.White)
                lab.BackColor = Color.Green;
            else
                lab.BackColor = Color.White;
        }

        private void Start_auto_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("CHAY AUTO");
            // GO HOME
            if (Start == 0)
            {
                hanhtrinh = DT.RowCount - 1;
                if (hanhtrinh > 0)
                {

                    Start_auto.Text = "RUN";
                    Start_auto.BackColor = Color.LimeGreen;
                    Manual.Enabled = false;

                    Start = 1;
                    _Req_ORG = 1;
                    AUTO_RUN = 1;
                    _READY_RUN = 0;
                    time_spray = 0;
                }
                else
                    MessageBox.Show("Chưa có dữ liệu tọa độ");
                

            }
            else
            {
                Start_auto.Text = "STOP";
                Start_auto.BackColor = Color.Red;
                Manual.Enabled = true;
    
                Start = 0;
                AUTO_RUN = 0;
                _READY_RUN = 0;
                try
                {
                    MB.WriteSingleRegister(ADD_AM, _Manual);
                }
                catch
                {
                    MessageBox.Show("Auto manual");
                }
                _AM = _Manual;
                Automanual.Text = "Manual";


            }

        }

        private void Pause_Click(object sender, EventArgs e)
        {
            DT.DataSource = null;
            if (_PAUSE == 0)
            {
                _PAUSE = 1;
                Pause.Text = "Pause";
                Pause.BackColor = Color.Yellow;
            }
            else
            {
                _PAUSE = 0;
                Pause.Text = "Play";
                Pause.BackColor = Color.LimeGreen;
            }
        }
        private void CHECK_AM()
        {
            if (_AM == _Semi_auto)
            {
                Automanual.Text = "Semi_auto";
                Set_speed_manual.Enabled = false;
                Speed_manual.Enabled = false;

                //DK_XYZ.Hide();
                //DK_AUTO.Hide();
                //update_cood.Hide();
                //PUMP.Hide();
                //checkBox_Dispenser.Hide();

                DK_XYZ.Enabled = false;
                DK_AUTO.Enabled = false;
                update_cood.Enabled = false;
                PUMP.Enabled = false;
                checkBox_Dispenser.Enabled = false;
            }
            if (_AM == _Auto)
            {
                Automanual.Text = "Auto";
                Set_speed_manual.Enabled = false;
                Speed_manual.Enabled = false;
                //DK_XYZ.Hide();
                //DK_AUTO.Show();
                //update_cood.Hide();
                //PUMP.Hide();
                //checkBox_Dispenser.Show();

                DK_XYZ.Enabled = false;
                DK_AUTO.Enabled = true;
                update_cood.Enabled = false;
                PUMP.Enabled = false;
                checkBox_Dispenser.Enabled = true;
                //==================== Clear bitnmap=======================
                Bitmap tmp = new Bitmap(450, 450);
                TULA = tmp;
                MAP.BackgroundImage = TULA;
                //==========================================================
            }
            if (_AM == _Manual)
            {
                Automanual.Text = "Manual";
                Set_speed_manual.Enabled = true;
                Speed_manual.Enabled = true;
                //DK_XYZ.Show();
                //DK_AUTO.Hide();
                //update_cood.Show();
                //PUMP.Show();
                //checkBox_Dispenser.Hide();

                DK_XYZ.Enabled = true;
                DK_AUTO.Enabled = false;
                update_cood.Enabled = true;
                PUMP.Enabled = true;
                checkBox_Dispenser.Enabled = false;
            }

        }
        private void Automanual_Click(object sender, EventArgs e)
        {
            _AM++;
            if (_AM > 2)
                _AM = 0;
            CHECK_AM();

            try
            {
                MB.WriteSingleRegister(_Add_AutoManual, _AM);
            }
            catch
            {

            }
        }

        private void Select_Speed_Click(object sender, EventArgs e)
        {
            //if (Select_Speed.Text == "LOW SPEED")
            //{
            //    Select_Speed.Text = "HIGH SPEED";
            //    _SPEED = SPEED_HIGH;
            //}
            //else
            //{
            //    Select_Speed.Text = "LOW SPEED";
            //    _SPEED = SPEED_LOW;
            //}



        }


        private void Add_Click(object sender, EventArgs e)
        {
            CheckBox CB = new CheckBox();
            CB.Checked = false;
            CheckBox CB_Dp = new CheckBox();
            CB_Dp.Checked = true;


            ID_HANHTRINH++;
            ID_PL.Text = ID_HANHTRINH.ToString();

            if(TYPE_HANHTRINH == POINT)
            {
                if(ID_PL.Text != "" && X1.Text != "" && Y1.Text != "" && Z1.Text != "")
                {

                    table.Rows.Add(ID_PL.Text, Type_PL.Text, X1.Text, Y1.Text, Z1.Text, X1.Text, Y1.Text, Z1.Text, "0", "0", CB.CheckState, CB_Dp.CheckState);
                    DT.DataSource = table;
                }
                else
                { 
                    if(ID_PL.Text == "" ) MessageBox.Show("Chưa nhập ID");
                    if(X1.Text == "" ) MessageBox.Show("Chưa nhập X1");
                    if(Y1.Text == "" ) MessageBox.Show("Chưa nhập Y1");
                    if(Z1.Text == "") MessageBox.Show("Chưa nhập Z1");
                }

            }
            if(TYPE_HANHTRINH == LINE)
            {
                if (ID_PL.Text != "" && X1.Text != "" && Y1.Text != "" && Z1.Text != "" && X2.Text != "" && Y2.Text != "" && Z2.Text != "")
                {
                    //ID_HANHTRINH++;
                    //ID_PL.Text = ID_HANHTRINH.ToString();
                    table.Rows.Add(ID_PL.Text, Type_PL.Text, X1.Text, Y1.Text, Z1.Text, X2.Text, Y2.Text, Z2.Text, "0", Speed_manual.Text, CB.CheckState, CB_Dp.CheckState);
                    DT.DataSource = table;
                }
                else
                {
                    if (ID_PL.Text == "") MessageBox.Show("Chưa nhập ID");
                    if (X1.Text == "") MessageBox.Show("Chưa nhập X1");
                    if (Y1.Text == "") MessageBox.Show("Chưa nhập Y1");
                    if (Z1.Text == "") MessageBox.Show("Chưa nhập Z1");
                    if (X2.Text == "") MessageBox.Show("Chưa nhập X2");
                    if (Y2.Text == "") MessageBox.Show("Chưa nhập Y2");
                    if (Z2.Text == "") MessageBox.Show("Chưa nhập Z2");
                }

            }
            if(TYPE_HANHTRINH == Rectangular_Array)
            {
                if (ID_PL.Text != "" && X1.Text != "" && Y1.Text != "" && Z1.Text != "" && X2.Text != "" && Y2.Text != "" && Z2.Text != "" && Number_step.Text != "")
                {
                    //ID_HANHTRINH++;
                    //ID_PL.Text = ID_HANHTRINH.ToString();
                    table.Rows.Add(ID_PL.Text, Type_PL.Text, X1.Text, Y1.Text, Z1.Text, X2.Text, Y2.Text, Z2.Text, Number_step.Text, Speed_manual.Text, CB.CheckState, CB_Dp.CheckState);
                    DT.DataSource = table;
                }
                else
                {
                    if (ID_PL.Text == "") MessageBox.Show("Chưa nhập ID");
                    if (X1.Text == "") MessageBox.Show("Chưa nhập X1");
                    if (Y1.Text == "") MessageBox.Show("Chưa nhập Y1");
                    if (Z1.Text == "") MessageBox.Show("Chưa nhập Z1");
                    if (X2.Text == "") MessageBox.Show("Chưa nhập X2");
                    if (Y2.Text == "") MessageBox.Show("Chưa nhập Y2");
                    if (Z2.Text == "") MessageBox.Show("Chưa nhập Z2");
                }

            }
            //================================= CHECK add SEQUENTIAL =================================
            if(check_Sequential.Checked == true)
            {
                X1.Text = X2.Text;
                Y1.Text = Y2.Text;
                Z1.Text = Z1.Text;
            }
                 
        }

        private void DT_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //DT.ContextMenuStrip = contextMenuStrip1;

        }

        private void Select_Speed_MouseDown(object sender, MouseEventArgs e)
        {
            i_SPEED++;
            if (i_SPEED > 3)
                i_SPEED = 1;

            if (i_SPEED == 1)
                Select_Speed.Text = "LOW SPEED";
            if (i_SPEED == 2)
                Select_Speed.Text = "MID SPEED";
            if (i_SPEED == 3)
                Select_Speed.Text = "HIGH SPEED";

            try
            {
                MB.WriteSingleRegister(_Add_SPEED, i_SPEED);
            }
            catch
            {

            }
        }

        private void T_ID_Click(object sender, EventArgs e)
        {

        }

        private void Type_PL_Click(object sender, EventArgs e)
        {
            TYPE_HANHTRINH++;
            if (TYPE_HANHTRINH > 3)
                TYPE_HANHTRINH = 1;
            if(TYPE_HANHTRINH == LINE)
            {
                Type_PL.Text = "Line";
                TOADO_2.Show();
                X2.Show();
                Y2.Show();
                Z2.Show();

                NB_steps.Hide();
                Number_step.Hide();
               
            }
            if(TYPE_HANHTRINH == POINT)
            {
                Type_PL.Text = "Point";
                TOADO_2.Hide();
                X2.Hide();
                Y2.Hide();
                Z2.Hide();
                NB_steps.Hide();
                Number_step.Hide();
            }
            if(TYPE_HANHTRINH == Rectangular_Array)
            {
                Type_PL.Text = "RECTANG";
                TOADO_2.Show();
                X2.Show();
                Y2.Show();
                Z2.Show();

                NB_steps.Show();
                Number_step.Show();

            }
        }

        private void TOADO_1_Click(object sender, EventArgs e)
        {
            if (choice_toado == 0 || choice_toado == CHOICE_TOADO_2)
            {
                choice_toado = CHOICE_TOADO_1;

                TOADO_1.ForeColor = Color.Green;
                TOADO_2.ForeColor = Color.Black;
                X1.Enabled = true;
                Y1.Enabled = true;
                Z1.Enabled = true;
                X2.Enabled = false;
                Y2.Enabled = false;
                Z2.Enabled = false;
            }
            else
            {
                choice_toado = 0;
                TOADO_1.ForeColor = Color.Black;
                X1.Enabled = false;
                Y1.Enabled = false;
                Z1.Enabled = false;
            }
        }

        private void TOADO_2_Click(object sender, EventArgs e)
        {
            if (choice_toado == 0 || choice_toado == CHOICE_TOADO_1)
            {
                choice_toado = CHOICE_TOADO_2;
                TOADO_2.ForeColor = Color.Green;
                TOADO_1.ForeColor = Color.Black;
                X2.Enabled = true;
                Y2.Enabled = true;
                Z2.Enabled = true;

                X1.Enabled = false;
                Y1.Enabled = false;
                Z1.Enabled = false;
            }
            else
            {
                choice_toado = 0;
                TOADO_2.ForeColor = Color.Black;
                X2.Enabled = false;
                Y2.Enabled = false;
                Z2.Enabled = false;
            }
        }
        void SAVE_DATA()
        {
            byte[] val16 = new byte[2];
            ushort val = 0;
            #region SAVE - SETTING
            ushort check_data_Error = 0;
            //------- Z MOVE --------------------
            val = ushort.Parse(Z_auto_move.Text);
            val16 = BitConverter.GetBytes(val);
            data[1] = val16[0];
            data[2] = val16[1];

            //------- TOA DO waiting -----------------
            val = ushort.Parse(Posw_x.Text);
            val16 = BitConverter.GetBytes(val);
            data[3] = val16[0];
            data[4] = val16[1];

            val = ushort.Parse(Posw_y.Text);
            val16 = BitConverter.GetBytes(val);
            data[5] = val16[0];
            data[6] = val16[1];

            val = ushort.Parse(Posw_z.Text);
            val16 = BitConverter.GetBytes(val);
            data[7] = val16[0];
            data[8] = val16[1];
            //============================================
            //val = ushort.Parse(Z_limit.Text);
            //val16 = BitConverter.GetBytes(val);
            //data[9] = val16[0];
            //data[10] = val16[1];
            //========================Disspray==================
            val = ushort.Parse(DSP_X.Text);
            val16 = BitConverter.GetBytes(val);
            data[11] = val16[0];
            data[12] = val16[1];

            val = ushort.Parse(DSP_Y.Text);
            val16 = BitConverter.GetBytes(val);
            data[13] = val16[0];
            data[14] = val16[1];

            val = ushort.Parse(DSP_Z.Text);
            val16 = BitConverter.GetBytes(val);
            data[15] = val16[0];
            data[16] = val16[1];
            //================== time drop =========================

            val = ushort.Parse(val_timedrop.Text);
            val16 = BitConverter.GetBytes(val);
            data[17] = val16[0];
            data[18] = val16[1];
            //==================== TOA DO OF1 ===================
            val = ushort.Parse(OF1_X.Text);
            val16 = BitConverter.GetBytes(val);
            data[19] = val16[0];
            data[20] = val16[1];

            val = ushort.Parse(OF1_Y.Text);
            val16 = BitConverter.GetBytes(val);
            data[21] = val16[0];
            data[22] = val16[1];
            //=================== TOA DO OFn ====================
            val = ushort.Parse(OFn_X.Text);
            val16 = BitConverter.GetBytes(val);
            data[23] = val16[0];
            data[24] = val16[1];

            val = ushort.Parse(OFn_Y.Text);
            val16 = BitConverter.GetBytes(val);
            data[25] = val16[0];
            data[26] = val16[1];
            //================== number copy ====================
            val = ushort.Parse(Number_copy.Text);
            val16 = BitConverter.GetBytes(val);
            data[27] = val16[0];
            data[28] = val16[1];
            if (val < 1)
            {
                MessageBox.Show("Number copy phai lon hon 1!");
                check_data_Error = 1;
            }
            //================== ID start copy ====================
            val = ushort.Parse(ID_start_copy.Text);
            val16 = BitConverter.GetBytes(val);
            data[29] = val16[0];
            data[30] = val16[1];
            if (val < 1 || val > DT.Rows.Count)
            {
                MessageBox.Show("ID start copy chua dung");
                check_data_Error = 1;
            }



            #endregion

            //============================================ SO HANH TRINH ==============================================
            ushort sohanhtrinh = (ushort)(DT.RowCount - 1);
            val16 = BitConverter.GetBytes(sohanhtrinh);
            data[99] = val16[0];
            data[100] = val16[1];
            // Add_START = 100;
            //=========================================== CÁC ĐIỂM HÀNH TRÌNH =========================================
            for (int i = 0; i < (DT.RowCount - 1); i++)
            {
                #region POINT
                if (String.Compare(DT.Rows[i].Cells[1].Value.ToString(), "Point") == 0)  //=================>> TOADO_1 = TOA_DO_2
                {
                    //--------------------- X1 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[2].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 1] = val16[0];
                    data[Add_START + i * SIZE_DATA + 2] = val16[1];
                    //--------------------- Y1 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[3].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 3] = val16[0];
                    data[Add_START + i * SIZE_DATA + 4] = val16[1];

                    //--------------------- Z1 ------------------------------------------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[4].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 5] = val16[0];
                    data[Add_START + i * SIZE_DATA + 6] = val16[1];

                    //--------------------- X2 ------------------------
                    val =  ushort.Parse(DT.Rows[i].Cells[5].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 7] = val16[0];
                    data[Add_START + i * SIZE_DATA + 8] = val16[1];
                    //--------------------- Y2 ------------------------
                    val =  ushort.Parse(DT.Rows[i].Cells[6].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 9] = val16[0];
                    data[Add_START + i * SIZE_DATA + 10] = val16[1];

                    //--------------------- Z2 ------------------------
                    val =  ushort.Parse(DT.Rows[i].Cells[7].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 11] = val16[0];
                    data[Add_START + i * SIZE_DATA + 12] = val16[1];

                    //---------------------- STEPS -------------------
                    val = 0;// ushort.Parse(DT.Rows[i].Cells[8].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 13] = val16[0];
                    //---------------------- speed -------------------
                    val = 0;// ushort.Parse(DT.Rows[i].Cells[9].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 14] = val16[0];
                    //-------------------- - LIFT up ----------------
                    if (bool.Parse(DT.Rows[i].Cells[10].Value.ToString()))
                    {
                        val = 1;
                    }
                    else
                        val = 0;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 15] = val16[0];
                    //================ DISPENSER ======================
                    if (bool.Parse(DT.Rows[i].Cells[11].Value.ToString()))
                    {
                        val = 1;
                    }
                    else
                        val = 0;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 16] = val16[0];
                    //=================================================
                }
                #endregion
                #region Line
                if (String.Compare(DT.Rows[i].Cells[1].Value.ToString(), "Line") == 0)  //============================>>>>>>>>>>>>>>>>>>
                {

                    //--------------------- X1 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[2].Value.ToString());
                    //val += 50;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 1] = val16[0];
                    data[Add_START + i * SIZE_DATA + 2] = val16[1];
                    //--------------------- Y1 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[3].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 3] = val16[0];
                    data[Add_START + i * SIZE_DATA + 4] = val16[1];

                    //--------------------- Z1 ------------------------------------------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[4].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 5] = val16[0];
                    data[Add_START + i * SIZE_DATA + 6] = val16[1];

                    //--------------------- X2 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[5].Value.ToString());
                    //val += 50;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 7] = val16[0];
                    data[Add_START + i * SIZE_DATA + 8] = val16[1];
                    //--------------------- Y2 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[6].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 9] = val16[0];
                    data[Add_START + i * SIZE_DATA + 10] = val16[1];

                    //--------------------- Z2 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[7].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 11] = val16[0];
                    data[Add_START + i * SIZE_DATA + 12] = val16[1];

                    //---------------------- STEPS -------------------
                    val = 0;// ushort.Parse(DT.Rows[i].Cells[8].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 13] = val16[0];
                    //---------------------- SPEED  -------------------
                    val = ushort.Parse(DT.Rows[i].Cells[9].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 14] = val16[0];
                    //----------------------- Lift up -----------------------
                    if (bool.Parse(DT.Rows[i].Cells[10].Value.ToString()))
                    {
                        val = 1;
                    }
                    else
                        val = 0;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 15] = val16[0];
                    //================ DISPENSER ======================
                    if (bool.Parse(DT.Rows[i].Cells[11].Value.ToString()))
                    {
                        val = 1;
                    }
                    else
                        val = 0;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 16] = val16[0];
                    //=================================================
                }
                #endregion
                #region RECTANG
                if (String.Compare(DT.Rows[i].Cells[1].Value.ToString(), "RECTANG") == 0)  //============================>>>>>>>>>>>>>>>>>>
                {

                    //--------------------- X1 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[2].Value.ToString());
                    //val += 50;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 1] = val16[0];
                    data[Add_START + i * SIZE_DATA + 2] = val16[1];
                    //--------------------- Y1 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[3].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 3] = val16[0];
                    data[Add_START + i * SIZE_DATA + 4] = val16[1];

                    //--------------------- Z1 ------------------------------------------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[4].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 5] = val16[0];
                    data[Add_START + i * SIZE_DATA + 6] = val16[1];

                    //--------------------- X2 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[5].Value.ToString());
                    //val += 50;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 7] = val16[0];
                    data[Add_START + i * SIZE_DATA + 8] = val16[1];
                    //--------------------- Y2 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[6].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 9] = val16[0];
                    data[Add_START + i * SIZE_DATA + 10] = val16[1];

                    //--------------------- Z2 ------------------------
                    val = ushort.Parse(DT.Rows[i].Cells[7].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 11] = val16[0];
                    data[Add_START + i * SIZE_DATA + 12] = val16[1];
                    //---------------------- STEPS -------------------
                    val = ushort.Parse(DT.Rows[i].Cells[8].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 13] = val16[0];
                    //---------------------- SPEED ------------------
                    val = ushort.Parse(DT.Rows[i].Cells[9].Value.ToString());
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 14] = val16[0];
                    //----------------------- lift up -------------------
                    if (bool.Parse(DT.Rows[i].Cells[10].Value.ToString()))
                    {
                        val = 1;
                    }
                    else
                        val = 0;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 15] = val16[0];
                    //================ DISPENSER ======================
                    if (bool.Parse(DT.Rows[i].Cells[11].Value.ToString()))
                    {
                        val = 1;
                    }
                    else
                        val = 0;
                    val16 = BitConverter.GetBytes(val);
                    data[Add_START + i * SIZE_DATA + 16] = val16[0];
                    //=================================================
                }
                #endregion

            }

            if (Project_running.Text != "" && check_data_Error == 0 && DT.DataSource != null)
            {
                File.WriteAllBytes(Project_running.Text + ".tula", data);
                MessageBox.Show("Save complete!");
            }
            else
                MessageBox.Show("Chua co Project running | Null ");


            check_data_Error = 0;
        }
        private void Save_data_Click(object sender, EventArgs e)
        {

            SAVE_DATA();
        }
        byte[] Data_toado;
        private void Call_data_Click(object sender, EventArgs e)
        {
  
            DT.DataSource = null;
            table.Clear();

            byte[] val16 = new byte[2];
            ushort val = 0;
            //=====================================================================================
            ushort X1_T, Y1_T, Z1_T, X2_T, Y2_T, Z2_T, STEP_T, SPEED_T, LIFT_UP_T, DISPENSER_T;
            int Type_T;
            //=============================  DATA IC_STACK ========================================//


            if (Project_running.Text != "")
            {

                Data_toado = File.ReadAllBytes(Project_running.Text + ".tula");


                #region DATA- SETTING
                //============================================  Z AUTO MOVE ============================================
                val16[0] = Data_toado[1];
                val16[1] = Data_toado[2];
                Zmove = BitConverter.ToUInt16(val16, 0);

                Z_auto_move.Text = Zmove.ToString();
                //=========================================== xyx WAITING ==============================================
                val16[0] = Data_toado[3];
                val16[1] = Data_toado[4];
                PW_X = BitConverter.ToUInt16(val16, 0);
                Posw_x.Text = PW_X.ToString();

                val16[0] = Data_toado[5];
                val16[1] = Data_toado[6];
                PW_Y = BitConverter.ToUInt16(val16, 0);
                Posw_y.Text = PW_Y.ToString();

                val16[0] = Data_toado[7];
                val16[1] = Data_toado[8];
                PW_Z = BitConverter.ToUInt16(val16, 0);
                Posw_z.Text = PW_Z.ToString();

                //==================================== limit z ==============================================
                //val16[0] = Data_toado[9];
                //val16[1] = Data_toado[10];
                //Zlimit = BitConverter.ToUInt16(val16, 0);
                //Z_limit.Text = Zlimit.ToString();
                //LIMIT_Z.Text = Zlimit.ToString();
                //================================== Disspray ================================
                val16[0] = Data_toado[11];
                val16[1] = Data_toado[12];
                Dsp_X = BitConverter.ToUInt16(val16, 0);
                DSP_X.Text = Dsp_X.ToString();

                val16[0] = Data_toado[13];
                val16[1] = Data_toado[14];
                Dsp_Y = BitConverter.ToUInt16(val16, 0);
                DSP_Y.Text = Dsp_Y.ToString();

                val16[0] = Data_toado[15];
                val16[1] = Data_toado[16];
                Dsp_Z = BitConverter.ToUInt16(val16, 0);
                DSP_Z.Text = Dsp_Z.ToString();
                //======================== time drop =========================================
                val16[0] = Data_toado[17];
                val16[1] = Data_toado[18];
                Time_Drop = BitConverter.ToUInt16(val16, 0);
                val_timedrop.Text = Time_Drop.ToString();
                //======================= OF1 ============================================
                val16[0] = Data_toado[19];
                val16[1] = Data_toado[20];
                OF1x = BitConverter.ToUInt16(val16, 0);
                OF1_X.Text = OF1x.ToString();

                val16[0] = Data_toado[21];
                val16[1] = Data_toado[22];
                OF1y = BitConverter.ToUInt16(val16, 0);
                OF1_Y.Text = OF1y.ToString();
                //====================== OFn =============================================
                val16[0] = Data_toado[23];
                val16[1] = Data_toado[24];
                OFnx = BitConverter.ToUInt16(val16, 0);
                OFn_X.Text = OFnx.ToString();

                val16[0] = Data_toado[25];
                val16[1] = Data_toado[26];
                OFny = BitConverter.ToUInt16(val16, 0);
                OFn_Y.Text = OFny.ToString();
                //===================== number copy========================================
                val16[0] = Data_toado[27];
                val16[1] = Data_toado[28];
                Numbercopy = BitConverter.ToUInt16(val16, 0);
                Number_copy.Text = Numbercopy.ToString();
                //===================== ID START copy========================================
                val16[0] = Data_toado[29];
                val16[1] = Data_toado[30];
                Id_start_copy = BitConverter.ToUInt16(val16, 0);
                ID_start_copy.Text = Id_start_copy.ToString();

                if (Numbercopy == 1)
                {
                    deX = deY = 0;
                }
                else if (Numbercopy > 1)
                {
                    //deX = Math.Abs(OFnx - OF1x) / (Numbercopy - 1);
                    //deY = Math.Abs(OFny - OF1y) / (Numbercopy - 1);

                    deX = Math.Abs(OFnx - OF1x);
                    deY = Math.Abs(OFny - OF1y);
                    textBox6.Text = deX.ToString();
                    textBox7.Text = deY.ToString();

                }
                #endregion
                #region DATA- HANH TRINH
                //----------------------------- SO HANH TRINH --------------------------------//
                val16[0] = Data_toado[99];
                val16[1] = Data_toado[100];
                ushort sohanhtrinh = BitConverter.ToUInt16(val16, 0);
                MessageBox.Show(sohanhtrinh.ToString());
                //============================ TOA DO CAC DIEM ================================//
                for (int i = 0; i < sohanhtrinh; i++)
                {

                    //================== X1 ======================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 1];
                    val16[1] = Data_toado[Add_START + i * SIZE_DATA + 2];
                    X1_T = BitConverter.ToUInt16(val16, 0);
                    X1_spray[i + 1] = X1_T;
                    //================= Y1 ======================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 3];
                    val16[1] = Data_toado[Add_START + i * SIZE_DATA + 4];
                    Y1_T = BitConverter.ToUInt16(val16, 0);
                    Y1_spray[i + 1] = Y1_T;
                    //================= Z1 =======================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 5];
                    val16[1] = Data_toado[Add_START + i * SIZE_DATA + 6];
                    Z1_T = BitConverter.ToUInt16(val16, 0);
                    Z1_spray[i + 1] = Z1_T;
                    //================== X2 ======================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 7];
                    val16[1] = Data_toado[Add_START + i * SIZE_DATA + 8];
                    X2_T = BitConverter.ToUInt16(val16, 0);
                    X2_spray[i + 1] = X2_T;
                    //================= Y2 ======================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 9];
                    val16[1] = Data_toado[Add_START + i * SIZE_DATA + 10];
                    Y2_T = BitConverter.ToUInt16(val16, 0);
                    Y2_spray[i + 1] = Y2_T;
                    //================= Z2 =======================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 11];
                    val16[1] = Data_toado[Add_START + i * SIZE_DATA + 12];
                    Z2_T = BitConverter.ToUInt16(val16, 0);
                    Z2_spray[i + 1] = Z2_T;
                    //================ STEPS ========================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 13];
                    val16[1] = 0;
                    STEP_T = BitConverter.ToUInt16(val16, 0);
                    STEP_spray[i + 1] = STEP_T;
                    //================ SPEED ========================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 14];
                    val16[1] = 0;
                    SPEED_T = BitConverter.ToUInt16(val16, 0);
                    SPEED_spray[i + 1] = SPEED_T;
                    //=================== lift up================================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 15];
                    val16[1] = 0;
                    LIFT_UP_T = BitConverter.ToUInt16(val16, 0);
                    LIFT_UP[i + 1] = LIFT_UP_T;

                    CheckBox CB = new CheckBox();
                    if (LIFT_UP_T == 1)
                        CB.Checked = true;
                    else
                        CB.Checked = false;
                    //==================== DISPENSER=========================
                    val16[0] = Data_toado[Add_START + i * SIZE_DATA + 16];
                    val16[1] = 0;
                    DISPENSER_T = BitConverter.ToUInt16(val16, 0);
                    DISPEN[i + 1] = DISPENSER_T;
                    CheckBox CB_DISPENSER = new CheckBox();
                    if (DISPENSER_T == 1)
                        CB_DISPENSER.Checked = true;
                    else
                        CB_DISPENSER.Checked = false;
                    //---------------------Type_spray-----------------------------

                    if ((X2_T == X1_T) && (Y2_T == Y1_T) && (Z2_T == Z1_T))
                    {
                        Type_T = POINT;
                        Type_spray[i + 1] = POINT;
                    }
                    else
                    {
                        if (STEP_spray[i + 1] == 0)
                        {
                            Type_T = LINE;
                            Type_spray[i + 1] = LINE;
                        }
                        else
                        {
                            Type_T = Rectangular_Array;
                            Type_spray[i + 1] = Rectangular_Array;
                            _stepY = (Math.Abs(Y2_T - Y1_T)) / STEP_T;
                        }
                    }


                    if (Type_T == POINT)
                        table.Rows.Add((i + 1), "Point", X1_T, Y1_T, Z1_T, "0", "0", "0", "0", CB.CheckState, CB_DISPENSER.CheckState);
                    if (Type_T == LINE)
                        table.Rows.Add((i + 1), "Line", X1_T.ToString(), Y1_T.ToString(), Z1_T.ToString(), X2_T.ToString(), Y2_T.ToString(), Z2_T.ToString(), "0", SPEED_T.ToString(), CB.CheckState, CB_DISPENSER.CheckState);
                    if (Type_T == Rectangular_Array)
                    {
                        table.Rows.Add((i + 1), "RECTANG", X1_T.ToString(), Y1_T.ToString(), Z1_T.ToString(), X2_T.ToString(), Y2_T.ToString(), Z2_T.ToString(), STEP_T.ToString(), SPEED_T.ToString(), CB.CheckState, CB_DISPENSER.CheckState);
                    }

                }
                DT.DataSource = table;
                ID_HANHTRINH = DT.RowCount - 1;
                ID_PL.Text = ID_HANHTRINH.ToString();

                Color_datagrid();
                #endregion
            }
            else
                MessageBox.Show("Chua co project running");

        }
        void Color_datagrid()
        {
            for(int i = 0;i < DT.Rows.Count - 1;i++)
            {
                if (String.Compare(DT.Rows[i].Cells[1].Value.ToString(), "Point") == 0) 
                    DT.Rows[i].DefaultCellStyle.BackColor = Color.Orange; 
                if (String.Compare(DT.Rows[i].Cells[1].Value.ToString(), "Line") == 0)
                    DT.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                if (String.Compare(DT.Rows[i].Cells[1].Value.ToString(), "RECTANG") == 0)
                    DT.Rows[i].DefaultCellStyle.BackColor = Color.Yellow; 
            }
        }
        private void Delete_Click(object sender, EventArgs e)
        {
            DT.DataSource = null;
            table.Clear();
            SO_HANHTRINH = 0;


        }

        private void Next_Click(object sender, EventArgs e)
        {
            hanhtrinh_so = 1;
            REQ_DISSPRAY = 1;

            Bitmap tmp = new Bitmap(450, 450);
            TULA = tmp;
            MAP.BackgroundImage = TULA;
  
        }

        private void Y_UP_Click(object sender, EventArgs e)
        {

        }


        private void tocdo_truc_Scroll(object sender, EventArgs e)
        {
            textBox1.Text = tocdo_truc.Value.ToString();
            MB.WriteSingleRegister(14, tocdo_truc.Value);
        }

        private void Save_Zmove_Click(object sender, EventArgs e)
        {

            SAVE_DATA();

        }

        private void Z_auto_label_Click(object sender, EventArgs e)
        {
            Z_auto_move.Text = _coodinateZ.ToString();
            MessageBox.Show("Cập nhật Z auto move = " + _coodinateZ.ToString());
        }

        private void Waiting_pos_label_Click(object sender, EventArgs e)
        {
            Posw_x.Text = _coodinateX.ToString();
            Posw_y.Text = _coodinateY.ToString();
            Posw_z.Text = _coodinateZ.ToString();
            MessageBox.Show("Cập nhật XYZ waiting pos: " + _coodinateX.ToString() + "," + _coodinateY.ToString() + "," + _coodinateZ.ToString());
        }

        private void Dowload_data_but_Click(object sender, EventArgs e)
        {
            Dowload_data_but.Enabled = false;
            //========================== TAT TIMER1 ==============================================
            timer1.Enabled = false;
            //========================== YEU CAU DOWLOAD DATA ====================================
            MB.WriteSingleRegister(ADD_REQ_DOWLOAD, 1);
            Thread.Sleep(200);
            byte[] val16 = new byte[2];
            //=====================================================================================
            ushort X1_T, Y1_T, Z1_T, X2_T, Y2_T, Z2_T, STEP_T, SPEED_T, LIFT_UP_T, DISPENSER_T;
            int Type_T;
            //=============================  DATA control ========================================//
            byte[] Data_toado = File.ReadAllBytes(Project_running.Text + ".tula");

            //----------------------------- SO HANH TRINH --------------------------------//
            val16[0] = Data_toado[99];
            val16[1] = Data_toado[100];
            ushort sohanhtrinh = BitConverter.ToUInt16(val16, 0);
            MB.WriteSingleRegister(ADD_SOHANHTIRNH_DOWLOAD, sohanhtrinh);

            Loaddata_progress.Value = 10;
            //============================ TOA DO CAC DIEM ================================//
            for (int i = 0; i < sohanhtrinh; i++)
            {

                //================== X1 ======================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 1];
                val16[1] = Data_toado[Add_START + i * SIZE_DATA + 2];
                X1_T = BitConverter.ToUInt16(val16, 0);
                X1_spray[i + 1] = X1_T;
                //================= Y1 ======================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 3];
                val16[1] = Data_toado[Add_START + i * SIZE_DATA + 4];
                Y1_T = BitConverter.ToUInt16(val16, 0);
                Y1_spray[i + 1] = Y1_T;
                //================= Z1 =======================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 5];
                val16[1] = Data_toado[Add_START + i * SIZE_DATA + 6];
                Z1_T = BitConverter.ToUInt16(val16, 0);
                Z1_spray[i + 1] = Z1_T;
                //================== X2 ======================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 7];
                val16[1] = Data_toado[Add_START + i * SIZE_DATA + 8];
                X2_T = BitConverter.ToUInt16(val16, 0);
                X2_spray[i + 1] = X2_T;
                //================= Y2 ======================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 9];
                val16[1] = Data_toado[Add_START + i * SIZE_DATA + 10];
                Y2_T = BitConverter.ToUInt16(val16, 0);
                Y2_spray[i + 1] = Y2_T;
                //================= Z2 =======================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 11];
                val16[1] = Data_toado[Add_START + i * SIZE_DATA + 12];
                Z2_T = BitConverter.ToUInt16(val16, 0);
                Z2_spray[i + 1] = Z2_T;
                //================ STEPS ========================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 13];
                val16[1] = 0;
                STEP_T = BitConverter.ToUInt16(val16, 0);
                STEP_spray[i + 1] = STEP_T;
                //================= SPEED =======================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 14];
                val16[1] = 0;
                SPEED_T = BitConverter.ToUInt16(val16, 0);
                SPEED_spray[i + 1] = SPEED_T;
 
                //=================== lift up================================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 15];
                val16[1] = 0;
                LIFT_UP_T = BitConverter.ToUInt16(val16, 0);
                LIFT_UP[i + 1] = LIFT_UP_T;
                //==================== DISPENSER=========================
                val16[0] = Data_toado[Add_START + i * SIZE_DATA + 16];
                val16[1] = 0;
                DISPENSER_T = BitConverter.ToUInt16(val16, 0);
                DISPEN[i + 1] = DISPENSER_T;
           
                //========================================================
                if ((X1_T == X2_T) && (Y1_T == Y2_T) && (Z1_T == Z2_T))
                {
                    Type_T = POINT;
                    Type_spray[i + 1] = POINT;
                }
                else
                {
                    if (STEP_spray[i + 1] == 0)
                    {
                        Type_T = LINE;
                        Type_spray[i + 1] = LINE;
                    }
                    else
                    {
                        Type_T = Rectangular_Array;
                        Type_spray[i + 1] = Rectangular_Array;
                       // _stepY = (Math.Abs(Y2_T - Y1_T)) / STEP_T;
                    }
                }

                MB.WriteSingleRegister(ADD_hanhtrinh_so, i + 1);
                MB.WriteSingleRegister(ADD_TYPE_HANHTRINH, Type_T);
                MB.WriteSingleRegister(ADD_X1, X1_T);
                MB.WriteSingleRegister(ADD_Y1, Y1_T);
                MB.WriteSingleRegister(ADD_Z1, Z1_T);

                MB.WriteSingleRegister(ADD_X2, X2_T);
                MB.WriteSingleRegister(ADD_Y2, Y2_T);
                MB.WriteSingleRegister(ADD_Z2, Z2_T);

                MB.WriteSingleRegister(ADD_STEPS, STEP_T);
                MB.WriteSingleRegister(ADD_SPEED, SPEED_T);
                MB.WriteSingleRegister(ADD_LIFT_UP, LIFT_UP_T);
                MB.WriteSingleRegister(ADD_EN_DISPENSER, DISPENSER_T);
            }
            Loaddata_progress.Value = 90;
            MB.WriteSingleRegister(ADD_Z_MOVE, Zmove);
            MB.WriteSingleRegister(ADD_XWP, PW_X);
            MB.WriteSingleRegister(ADD_YWP, PW_Y);
            MB.WriteSingleRegister(ADD_ZWP, PW_Z);

            MB.WriteSingleRegister(ADD_XDSP, Dsp_X);
            MB.WriteSingleRegister(ADD_YDSP, Dsp_Y);
            MB.WriteSingleRegister(ADD_ZDSP, Dsp_Z);

            MB.WriteSingleRegister(ADD_TIME_DROP, Time_Drop);

            MB.WriteSingleRegister(ADD_deX, deX);
            MB.WriteSingleRegister(ADD_deY, deY);
            MB.WriteSingleRegister(ADD_numcopy_array, Numbercopy);
            MB.WriteSingleRegister(ADD_ID_start_copy, Id_start_copy);

            MessageBox.Show("DOWLOAD COMPLETE!");
            // ========================= BAO CAO DA TRUYEN FILE XONG  =================================
            MB.WriteSingleRegister(ADD_REPORT_FILE_LOAD_OK, 1);

            //========================== XOA YEU CAU DOWLOAD DU LIEU ===================================
            Thread.Sleep(5000);
            MB.WriteSingleRegister(ADD_REPORT_FILE_LOAD_OK, 0);
            MB.WriteSingleRegister(ADD_REQ_DOWLOAD, 0);
            Loaddata_progress.Value = 100;
            timer1.Enabled = true;
            Dowload_data_but.Enabled = true;
           // Req_GOHOME();

        }

        private void Set_speed_manual_Scroll(object sender, EventArgs e)
        {
            if(_AM == _Manual)
            {
                Speed_manual.Text = Set_speed_manual.Value.ToString();

            }
        }

        private Bitmap MergedBitmaps(Bitmap bmp1, Bitmap bmp2)
        {
            Bitmap result = new Bitmap(Math.Max(bmp1.Width, bmp2.Width),
                                       Math.Max(bmp1.Height, bmp2.Height));
            using (Graphics g = Graphics.FromImage(result))
            {
                g.DrawImage(bmp2, Point.Empty);
                g.DrawImage(bmp1, Point.Empty);
            }
            return result;
        }

        private void checkBox_Dispenser_CheckedChanged(object sender, EventArgs e)
        {
            if (_dispenser == 1)
                _dispenser = 0;
            else
                _dispenser = 1;
                     
            textBox2.Text = _dispenser.ToString();
        }

        private void PUMP_MouseUp(object sender, MouseEventArgs e)
        {
            if (_Status_SPRAY == Spray_OFF)
            {
                _Status_SPRAY = Spray_ON;
                PUMP.BackColor = Color.Green;
            }
            else
            {
                _Status_SPRAY = Spray_OFF;
                PUMP.BackColor = Color.Red;
            }
            try
            {
                MB.WriteSingleRegister(ADD_SPRAY, _Status_SPRAY);
            }
            catch
            {

            }
        }

        //private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        //{
        //    System.Diagnostics.Process.Start(e.Link.LinkData.ToString());
        //}

        private void Disspray_Click(object sender, EventArgs e)
        {
            DSP_X.Text = _coodinateX.ToString();
            DSP_Y.Text = _coodinateY.ToString();
            DSP_Z.Text = _coodinateZ.ToString();
            MessageBox.Show("Cập nhật XYZ Disspray: " + _coodinateX.ToString() + "," + _coodinateY.ToString() + "," + _coodinateZ.ToString());
        }

        private void Lab_OF1_Click(object sender, EventArgs e)
        {
            OF1_X.Text = _coodinateX.ToString();
            OF1_Y.Text = _coodinateY.ToString();
       
            MessageBox.Show("Cập nhật OF1: " + _coodinateX.ToString() + "," + _coodinateY.ToString());
        }

        private void label15_Click(object sender, EventArgs e)
        {
            OFn_X.Text = _coodinateX.ToString();
            OFn_Y.Text = _coodinateY.ToString();

            MessageBox.Show("Cập nhật OFn: " + _coodinateX.ToString() + "," + _coodinateY.ToString());
        }


        private void CopyClipboard()
        {
            DataObject d = DT.GetClipboardContent();

            if (d != null)
                Clipboard.SetDataObject(d);
        }


        /// <summary>
        /// This will be moved to the util class so it can service any paste into a DGV
        /// </summary>
        private void PasteClipboard()
        {
            //Show Error if no cell is selected
            if (DT.SelectedCells.Count == 0)
            {
                MessageBox.Show("Please select a cell", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Get the satring Cell
            DataGridViewCell startCell = GetStartCell(DT);
            //Get the clipboard value in a dictionary
            Dictionary<int, Dictionary<int, string>> cbValue = ClipBoardValues(Clipboard.GetText());

            int iRowIndex = startCell.RowIndex;
            foreach (int rowKey in cbValue.Keys)
            {
                int iColIndex = startCell.ColumnIndex;
                foreach (int cellKey in cbValue[rowKey].Keys)
                {
                    //Check if the index is with in the limit
                    if (iColIndex <= DT.Columns.Count - 1 && iRowIndex <= DT.Rows.Count - 1)
                    {
                        DataGridViewCell cell = DT[iColIndex, iRowIndex];

                        //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                        if ((chkPasteToSelectedCells.Checked && cell.Selected) || (!chkPasteToSelectedCells.Checked))
                            cell.Value = cbValue[rowKey][cellKey];
                    }
                    iColIndex++;
                }
                iRowIndex++;
            }
        }
        private DataGridViewCell GetStartCell(DataGridView dgView)
        {
            //get the smallest row,column index
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }

            return dgView[colIndex, rowIndex];
        }
        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>> copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                //if an empty cell value copied, then set the dictionay with an empty string
                //else Set value to dictionary
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }
        private void DT_KeyDown(object sender, KeyEventArgs e)
        {
       

            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyClipboard();
                            break;

                        case Keys.V:
                            PasteClipboard();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Copy/paste operation failed. " + ex.Message, "Copy/Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyClipboard();
        }

        private void pasteCtrlVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboard();
        }

        private void Insert_Top_ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Insert_Top_Clipboard();
        }

        private void Insert_Top_Clipboard()
        {

            if (row_selected > 1)
            {
                CheckBox CB = new CheckBox();
                CB.Checked = true;
                CheckBox CB_Dp = new CheckBox();
                CB_Dp.Checked = true;


                DataRow dr = table.NewRow();
                dr[0] = 1;
                dr[1] = "Line";
                dr[2] = 2;
                dr[3] = 3;
                dr[4] = 4;
                dr[5] = 5;
                dr[6] = 6;
                dr[7] = 7;
                dr[8] = 0;
                dr[9] = 2;
                dr[10] = CB.CheckState;
                dr[11] = CB.CheckState;
                table.Rows.InsertAt(dr, row_selected - 1);
               // DT.DataSource = null;
                DT.DataSource = table;
                row_selected = 0;
            }
        }
        private void Insert_BOT_toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Insert_Bot_Clipboard();
        }

        private void Insert_Bot_Clipboard()
        {
            if (row_selected > 0)
            {
                CheckBox CB = new CheckBox();
                CB.Checked = true;
                CheckBox CB_Dp = new CheckBox();
                CB_Dp.Checked = true;
                DataRow dr = table.NewRow();
                dr[0] = 1;
                dr[1] = "Line";
                dr[2] = 2;
                dr[3] = 3;
                dr[4] = 4;
                dr[5] = 5;
                dr[6] = 6;
                dr[7] = 7;
                dr[8] = 0;
                dr[9] = 2;
                dr[10] = CB.CheckState;
                dr[11] = CB.CheckState;
                table.Rows.InsertAt(dr, row_selected);
                DT.DataSource = table;
                row_selected = 0;
            }
        }
        private void MoveTo_Click(object sender, EventArgs e)
        {
            MoveTo_Clipboard();
        }

        private void MoveTo_Clipboard()
        {
            DataGridViewRow dt = new DataGridViewRow();
            dt = DT.Rows[row_selected-1];
            if(cell_selected < 5)
                MessageBox.Show("Move to " + dt.Cells[2].Value.ToString() + "," + dt.Cells[3].Value.ToString() + "," + dt.Cells[4].Value.ToString());
            else
                MessageBox.Show("Move to " + dt.Cells[5].Value.ToString() + "," + dt.Cells[6].Value.ToString() + "," + dt.Cells[7].Value.ToString());

        }
        private void DT_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (DT.SelectedCells.Count > 0)
                DT.ContextMenuStrip = contextMenuStrip1;
         
        }
        //int i_choice = 0;
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    if (i_choice > 0)
        //        DT.Rows[i_choice-1].Selected = false;
        //    i_choice++;
        //    if (i_choice >= DT.RowCount) i_choice = DT.RowCount - 1;
        //    DT.Rows[i_choice-1].Selected = true;
        //}

        private void button2_Click_1(object sender, EventArgs e)
        {
            CheckBox CB = new CheckBox();
            CB.Checked = true;
            CheckBox CB_Dp = new CheckBox();
            CB_Dp.Checked = true;

            //this.DT.Rows.Insert (3, 1);//, "abc", 2, 3, 4, 5, 6, 7, 8,9, CB.CheckState, CB_Dp.CheckState);
            DataRow dr =table.NewRow();
            dr[0] = 1;
            dr[1] = "Line";
            dr[2] = 2;
            dr[3] = 3;
            dr[4] = 4;
            dr[5] = 5;
            dr[6] = 6;
            dr[7] = 7;
            dr[8] = 0;
            dr[9] = 2;
            dr[10] = CB.CheckState;
            dr[11] = CB.CheckState;
            table.Rows.InsertAt(dr, 1);
            DT.DataSource = table;
            //DT.RowCount++;
            //this.DT.Rows.Add(3, 1);//, "abc", 2, 3, 4, 5, 6, 7, 8,9, CB.CheckState, CB_Dp.CheckState);
           // table.Rows.Insert(3, "123", "abc", "1", "2", "3", "0", "0", "0", "0", CB.CheckState, CB_Dp.CheckState);
            //table.Rows.Add(ID_PL.Text, Type_PL.Text, X1.Text, Y1.Text, Z1.Text, "0", "0", "0", "0", CB.CheckState, CB_Dp.CheckState);
        }

        private void DT_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show((e.ColumnIndex +1).ToString());
            row_selected = e.RowIndex + 1;
            cell_selected = e.ColumnIndex + 1;
        }

        private void Creat_File_Click_1(object sender, EventArgs e)
        {
            byte[] a = new byte[10];
            File.WriteAllBytes(File_name.Text + ".tula", a);
        }

        private void Delete_File_Click_1(object sender, EventArgs e)
        {
            File.Delete(File_name.Text + ".tula");
        }

        private void Remove_Click(object sender, EventArgs e)
        {
            File.Delete(List_file.Items[List_file.SelectedIndex].ToString() + ".tula");
            List_file.Items.RemoveAt(List_file.SelectedIndex);
            Write_name_project();
        }

        private void Import_project_Click(object sender, EventArgs e)
        {
            if (List_file.SelectedItem != null)
            {
                Project_running.Text = List_file.Items[List_file.SelectedIndex].ToString();
                MessageBox.Show("Selected: " + Project_running.Text);
            }
            else
                MessageBox.Show("Chua chon Project");
        }



        private void Add_file_Click(object sender, EventArgs e)
        {
            if (File_name.Text != "")
            {
                List_file.Items.Add(File_name.Text);

                Write_name_project();
                File_name.Focus();
                File_name.Clear();
            }
            else
            {
                MessageBox.Show("Chua nhap ten File");
                File_name.Focus();
            }
        }
        void Write_name_project()
        {
            
            using (StreamWriter SW = new StreamWriter("Name_project.txt"))
            {

                foreach (var lbt in List_file.Items)
                {
                    SW.WriteLine(lbt.ToString());
                }

            }
        }

        private void linkLabel2_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString());
        }

        private void checkBox_ReqCalib_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_ReqCalib.Checked == true)
            {
                NF1_X.Text = "0";
                NF1_Y.Text = "0";
                NFn_X.Text = "0";
                NFn_Y.Text = "0";
            }
            Calib_Zig.Enabled = !Calib_Zig.Enabled;
        }

        private void Lab_NF1_Click(object sender, EventArgs e)
        {
            NF1_X.Text = _coodinateX.ToString();
            NF1_Y.Text = _coodinateY.ToString();

            MessageBox.Show("Cập nhật NF1: " + _coodinateX.ToString() + "," + _coodinateY.ToString());
        }

        private void Lab_NFn_Click(object sender, EventArgs e)
        {
            NFn_X.Text = _coodinateX.ToString();
            NFn_Y.Text = _coodinateY.ToString();

            MessageBox.Show("Cập nhật NFn: " + _coodinateX.ToString() + "," + _coodinateY.ToString());
        }

        private void Calibration_cood()
        {
            byte[] val16 = new byte[2];
            ushort val = 0;
            for (int i = 0; i < (DT.RowCount - 1); i++)
            {
                //--------------------- X1 ------------------------
                val = Convert.ToUInt16(int.Parse(DT.Rows[i].Cells[2].Value.ToString()) + Co_Calib.dX);
                DT.Rows[i].Cells[2].Value = val;
                //val16 = BitConverter.GetBytes(val);
                //data[Add_START + i * SIZE_DATA + 1] = val16[0]; 
                //data[Add_START + i * SIZE_DATA + 2] = val16[1];
                //--------------------- Y1 ------------------------
                //val = ushort.Parse(DT.Rows[i].Cells[3].Value.ToString());
                val = Convert.ToUInt16(int.Parse(DT.Rows[i].Cells[3].Value.ToString()) + Co_Calib.dY);
                DT.Rows[i].Cells[3].Value = val;
                //val16 = BitConverter.GetBytes(val);
                //data[Add_START + i * SIZE_DATA + 3] = val16[0];
                //data[Add_START + i * SIZE_DATA + 4] = val16[1];

                //--------------------- Z1 ------------------------------------------------------------
                //val = ushort.Parse(DT.Rows[i].Cells[4].Value.ToString());
                //val16 = BitConverter.GetBytes(val);
                //data[Add_START + i * SIZE_DATA + 5] = val16[0];
                //data[Add_START + i * SIZE_DATA + 6] = val16[1];

                //--------------------- X2 ------------------------
                //val = ushort.Parse(DT.Rows[i].Cells[5].Value.ToString());
                val = Convert.ToUInt16(int.Parse(DT.Rows[i].Cells[5].Value.ToString()) + Co_Calib.dX);
                DT.Rows[i].Cells[5].Value = val;
                //val16 = BitConverter.GetBytes(val);
                //data[Add_START + i * SIZE_DATA + 7] = val16[0];
                //data[Add_START + i * SIZE_DATA + 8] = val16[1];
                //--------------------- Y2 ------------------------
                //val = ushort.Parse(DT.Rows[i].Cells[6].Value.ToString());
                val = Convert.ToUInt16(int.Parse(DT.Rows[i].Cells[6].Value.ToString()) + Co_Calib.dY);
                DT.Rows[i].Cells[6].Value = val;
                //val16 = BitConverter.GetBytes(val);
                //data[Add_START + i * SIZE_DATA + 9] = val16[0];
                //data[Add_START + i * SIZE_DATA + 10] = val16[1];

                //--------------------- Z2 ------------------------
                //val = ushort.Parse(DT.Rows[i].Cells[7].Value.ToString());
                //val16 = BitConverter.GetBytes(val);
                //data[Add_START + i * SIZE_DATA + 11] = val16[0];
                //data[Add_START + i * SIZE_DATA + 12] = val16[1];             
            }
  
        }
        private void BUT_Calib_update_Click(object sender, EventArgs e)
        {
            ushort val = 0;
            try
            {

                Co_Calib.dX = int.Parse(NF1_X.Text) - int.Parse(OF1_X.Text);
                Co_Calib.dY = int.Parse(NF1_Y.Text) - int.Parse(OF1_Y.Text);

                OF1_X.Text = NF1_X.Text;
                OF1_Y.Text = NF1_Y.Text;

                OFn_X.Text = NFn_X.Text;
                OFn_Y.Text = NFn_Y.Text;

                Calibration_cood();

                MessageBox.Show("Calibration Compelete: " + "dX:" + Co_Calib.dX.ToString() + " dY:" + Co_Calib.dY.ToString());
          
                checkBox_ReqCalib.Checked = false;
            }
            catch
            {
                MessageBox.Show("OF1-OFn Null");
            }
        }

    }

}

