using Microsoft.Win32;
using PCSC;
using PCSC.Iso7816;
using PCSC.Reactive;
using PCSC.Reactive.Events;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendIdTessera
{
    public partial class FormSendIdTessera : Form
    {
        [DllImport("kernel32.dll")]
        static extern void Sleep(int dwMilliseconds);

        [DllImport("MasterRD.dll")]
        static extern int lib_ver(ref uint pVer);

        [DllImport("MasterRD.dll")]
        static extern int rf_init_com(int port, int baud);

        [DllImport("MasterRD.dll")]
        static extern int rf_ClosePort();

        [DllImport("MasterRD.dll")]
        static extern int rf_antenna_sta(short icdev, byte mode);

        [DllImport("MasterRD.dll")]
        static extern int rf_init_type(short icdev, byte type);

        [DllImport("MasterRD.dll")]
        static extern int rf_request(short icdev, byte mode, ref ushort pTagType);

        [DllImport("MasterRD.dll")]
        static extern int rf_anticoll(short icdev, byte bcnt, IntPtr pSnr, ref byte pRLength);

        [DllImport("MasterRD.dll")]
        static extern int rf_select(short icdev, IntPtr pSnr, byte srcLen, ref sbyte Size);

        [DllImport("MasterRD.dll")]
        static extern int rf_halt(short icdev);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_authentication2(short icdev, byte mode, byte secnr, IntPtr key);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_initval(short icdev, byte adr, Int32 value);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_increment(short icdev, byte adr, Int32 value);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_decrement(short icdev, byte adr, Int32 value);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_readval(short icdev, byte adr, ref Int32 pValue);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_read(short icdev, byte adr, IntPtr pData, ref byte pLen);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_write(short icdev, byte adr, IntPtr pData);

        [DllImport("MasterRD.dll")]
        static extern int rf_beep(short icdev, int msec);

        [DllImport("MasterRD.dll")]
        static extern int rf_light(short icdev, int color);

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        private bool bConnectedDevice = false;
        private bool allowshowdisplay = false;
        private String carattereIniziale = "";
        private String carattereFinale = "";

        public FormSendIdTessera()
        {
            if (CheckInstance())
            {
                MessageBox.Show(this, "Un'altra istanza è già in esecuzione.", "Impossibile avviare", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(1);
                Application.Exit();
            }

            InitializeComponent();

            LeggiCaratteriSentinella();

            backgroundWorker.RunWorkerAsync();

        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(allowshowdisplay ? value : allowshowdisplay);
        }

        private void LeggiCaratteriSentinella()
        {
            this.carattereIniziale = "";
            this.carattereFinale = "";
        }

        private bool CheckInstance()
        {
            int np = Process.GetProcessesByName("SendIdTessera").Count<Process>();
            return np > 1;
        }

        private void SendToActiveWindow(String cardNumber)
        {
            if (cardNumber == String.Empty) return;
            cardNumber = carattereIniziale + cardNumber + carattereFinale + Environment.NewLine;
            IntPtr h = GetForegroundWindow();
            SendKeys.SendWait(cardNumber);
        }

        private void Connect()
        {
            IDisposable subscription = null;
            try
            {
                var readers = GetReaders();

                if (!readers.Any())
                {
                    Console.WriteLine("You need at least one connected smart card reader.");
                    Console.ReadKey();
                }

                var monitorFactory = MonitorFactory.Instance;

                subscription = monitorFactory
                    .CreateObservable(SCardScope.System, readers)
                    .Subscribe(OnNext, OnError);

                bConnectedDevice = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                bConnectedDevice = false;
                if(subscription != null)
                    subscription.Dispose();
            }
        }

        private string[] GetReaders()
        {
            var contextFactory = ContextFactory.Instance;
            using (var ctx = contextFactory.Establish(SCardScope.System))
            {
                return ctx.GetReaders();
            }
        }

        private void OnError(Exception exception)
        {
            Console.WriteLine("ERROR: {0}", exception.Message);
        }

        private void OnNext(MonitorEvent ev)
        {
            if (ev.GetType().ToString().Contains("CardInserted"))
            {
                String id = GetCardNumber();
                SendToActiveWindow(id);
            }
        }

        public String GetCardNumber()
        {
            try
            {
                var contextFactory = ContextFactory.Instance;
                using (var ctx = contextFactory.Establish(SCardScope.System))
                {
                    var readerNames = ctx.GetReaders();
                    var name = readerNames[0];

                    using (var isoReader = new IsoReader(
                        context: ctx,
                        readerName: name,
                        mode: SCardShareMode.Shared,
                        protocol: SCardProtocol.Any,
                        releaseContextOnDispose: false))
                    {
                        //Seleziono il master file
                        var apduMaster = new CommandApdu(IsoCase.Case3Short, isoReader.ActiveProtocol)
                        {
                            CLA = 0x00, // Class
                            Instruction = InstructionCode.SelectFile,
                            P1 = 0x00, // Parameter 1
                            P2 = 0x00, // Parameter 2
                            Data = new byte[] { 0x3F, 0x00 }
                        };

                        var responseMaster = isoReader.Transmit(apduMaster);

                        //Seleziono la cartella
                        var apduFolder = new CommandApdu(IsoCase.Case3Short, isoReader.ActiveProtocol)
                        {
                            CLA = 0x00, // Class
                            Instruction = InstructionCode.SelectFile,
                            P1 = 0x00, // Parameter 1
                            P2 = 0x00, // Parameter 2
                            Data = new byte[] { 0x10, 0x00 }
                        };

                        var responseFolder = isoReader.Transmit(apduFolder);

                        //Seleziono il file id tessera
                        var apduId = new CommandApdu(IsoCase.Case3Short, isoReader.ActiveProtocol)
                        {
                            CLA = 0x00, // Class
                            Instruction = InstructionCode.SelectFile,
                            P1 = 0x00, // Parameter 1
                            P2 = 0x00, // Parameter 2
                            Data = new byte[] { 0x10, 0x03 }
                        };

                        var responseId = isoReader.Transmit(apduId);

                        //Leggo il file id tessera
                        var apduRead = new CommandApdu(IsoCase.Case4Short, isoReader.ActiveProtocol)
                        {
                            CLA = 0x00, // Class
                            Instruction = InstructionCode.ReadBinary,
                            P1 = 0x00, // Parameter 1
                            P2 = 0x00, // Parameter 2
                            Data = new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }
                        };
                        var responseRead = isoReader.Transmit(apduRead);

                        if (!responseRead.HasData)
                        {
                            return "";
                        }
                        else
                        {
                            var data = responseRead.GetData();
                            return System.Text.Encoding.ASCII.GetString(data);
                        }
                    }
                }
            }
            catch(Exception)
            {
                return "";
            }
        }

        private void toolStripEsci_Click(object sender, EventArgs e)
        {
            Environment.Exit(1);
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            notifyIcon.Icon = Resources.logo_ko;
            notifyIcon.Visible = true;
            notifyIcon.ShowBalloonTip(500, "SendIdTessera", "La app sta funzionando in background", ToolTipIcon.Info);

            do
            {
                while (!bConnectedDevice)
                    Connect();
            } while (true);
        }

        private void FormSendKeys_Shown(object sender, EventArgs e)
        {
            this.Hide();
            this.Visible = false;
        }
    }
}
