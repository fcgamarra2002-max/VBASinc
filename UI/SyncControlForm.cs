using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using VBIDE = Microsoft.Vbe.Interop;

namespace VBASinc.UI
{
    /// <summary>
    /// COM-visible controller that shows the mission-critical sync UI.
    /// VBA only needs: CreateObject("VBASinc.SyncController").ShowUI ThisWorkbook.VBProject
    /// </summary>
    [ComVisible(true)]
    [Guid("F1E2D3C4-B5A6-7890-1234-56789ABCDEF0")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("VBASinc.SyncController")]
    public class SyncController : ISyncController
    {
        private const string DEFAULT_EXTERNAL_PATH = @"C:\src_vba";

        /// <summary>
        /// Shows the sync control panel UI.
        /// </summary>
        /// <param name="vbaProject">VBProject COM object from Excel/VBA.</param>
        public void ShowUI(object vbaProject)
        {
            ShowUI(vbaProject, DEFAULT_EXTERNAL_PATH);
        }

        /// <summary>
        /// Shows the sync control panel UI with custom path.
        /// </summary>
        /// <param name="vbaProject">VBProject COM object.</param>
        /// <param name="externalPath">Path to external source folder.</param>
        public void ShowUI(object vbaProject, string externalPath)
        {
            var project = vbaProject as VBIDE.VBProject;
            if (project == null)
            {
                MessageBox.Show("Invalid VBProject object.", "VBASinc", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var form = new SyncControlForm(project, externalPath))
            {
                form.ShowDialog();
            }
        }
    }

    /// <summary>
    /// COM interface for SyncController.
    /// </summary>
    [ComVisible(true)]
    [Guid("E2D3C4B5-A678-9012-3456-789ABCDEF012")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISyncController
    {
        void ShowUI(object vbaProject);
        void ShowUI(object vbaProject, string externalPath);
    }

    /// <summary>
    /// Mission-Critical Sync Control Panel.
    /// Single-view, state-reflective UI with folder selection, export/import, logs, and change history.
    /// </summary>
    internal class SyncControlForm : Form
    {
        #region Fields

        private VBIDE.VBProject? _vbaProject;
        private string _externalPath;
        private readonly Host.AddInHost? _host;
        // Motor manual eliminado
        private UIState _state = UIState.Init;

        private readonly List<ChangeRecord> _changeHistory = new List<ChangeRecord>();

        // Controls - Path Selection
        private Label lblTitle = null!;
        private Label lblStatus = null!;
        private Panel pnlStatusIndicator = null!;
        private GroupBox grpPath = null!;
        private TextBox txtPath = null!;
        private Button btnBrowse = null!;
        private FileSystemWatcher? _uiWatcher;
        
        // Controls - Multi-Project
        private ComboBox cboProjects = null!;
        private Button btnRefreshProjects = null!;
        private PictureBox picProjectSyncStatus = null!;

        // Controls - Precheck
        private GroupBox grpPrecheck = null!;
        private Label lblInternal = null!, lblExternal = null!, lblExport = null!, lblImport = null!;
        private Label lblInternalVal = null!, lblExternalVal = null!, lblExportVal = null!, lblImportVal = null!;

        // Controls - Actions
        // Create controls removed
        private Button btnAutoSync = null!;
        // Interval controls removed

        // Controls - Certificate (Removed)
        private GroupBox grpCertificate = null!;
        private Label lblCertStatusVal = null!, lblCertIntegrityVal = null!, lblCertCleanupVal = null!, lblCertHashVal = null!;

        // Controls - Logs & History
        private TabControl tabLogs = null!;
        private TextBox txtLog = null!;
        private ListView lstChanges = null!;

        // Controls - Footer
        private Button btnClose = null!;
        private Label lblMessage = null!;
        private Label lblElapsed = null!;
        private ImageList imgActions = null!;



        public VBIDE.VBProject? VbaProject
        {
            get => _vbaProject;
            set
            {
                _vbaProject = value;
                if (_vbaProject != null) ExecutePrecheck();
            }
        }

        #endregion

        #region Constructor

        public SyncControlForm(VBIDE.VBProject? vbaProject, string externalPath, Host.AddInHost? host = null)
        {
            _vbaProject = vbaProject;
            _host = host;
            
            // Use path from host if available (has the saved/updated path)
            if (_host != null && !string.IsNullOrEmpty(_host.ExternalPath))
                _externalPath = _host.ExternalPath;
            else
                _externalPath = externalPath;
            
            InitializeComponent();

            // Refrescar al activar la ventana (cuando el usuario vuelve al panel)
            this.Activated += (s, e) => ExecutePrecheck();

            SetupUIWatcher();
        }

        private void SetupUIWatcher()
        {
            try
            {
                if (_uiWatcher != null)
                {
                    _uiWatcher.EnableRaisingEvents = false;
                    _uiWatcher.Dispose();
                }

                if (!string.IsNullOrEmpty(_externalPath) && Directory.Exists(_externalPath))
                {
                    _uiWatcher = new FileSystemWatcher(_externalPath);
                    _uiWatcher.IncludeSubdirectories = true;
                    _uiWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size;
                    _uiWatcher.Changed += (s, e) => this.BeginInvoke(new Action(ExecutePrecheck));
                    _uiWatcher.Created += (s, e) => this.BeginInvoke(new Action(ExecutePrecheck));
                    _uiWatcher.Deleted += (s, e) => this.BeginInvoke(new Action(ExecutePrecheck));
                    _uiWatcher.Renamed += (s, e) => this.BeginInvoke(new Action(ExecutePrecheck));
                    _uiWatcher.EnableRaisingEvents = true;
                }
            }
            catch { }
        }

        #endregion

        #region UI Initialization

        private void InitializeComponent()
        {
            // Form settings
            this.Text = "VBASinc - Panel de Control";
            this.Size = new Size(520, 800);
            try { this.Icon = new Icon(typeof(Connect).Assembly.GetManifestResourceStream("VBASinc.logo.ico")); } catch { }
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.Font = new Font("Segoe UI", 9F);

            this.ShowIcon = false;

            int y = 13;

            // ===== HEADER =====
            // Logo
            var picLogo = new PictureBox
            {
                Size = new Size(80, 80),
                Location = new Point(10, y - 15),
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = new Bitmap(typeof(Connect).Assembly.GetManifestResourceStream("VBASinc.logo.png"))
            };
            this.Controls.Add(picLogo);
            picLogo.SendToBack(); // Enviar al fondo para no tapar nada

            lblTitle = new Label
            {
                Text = "MOTOR DE SINCRONIZACIÓN VBA",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold), // Reducido a 12pt
                Location = new Point(100, y + 10),
                Size = new Size(330, 30), // Aumentado ancho
                ForeColor = Color.FromArgb(51, 51, 51),
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblTitle);



            var lblVersion = new Label
            {
                Text = "V 1.0.4",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Location = new Point(100, y + 40),
                Size = new Size(320, 18),
                ForeColor = Color.FromArgb(120, 120, 120),
                TextAlign = ContentAlignment.MiddleCenter // Centrado respecto a su ancho
            };
            this.Controls.Add(lblVersion);

            pnlStatusIndicator = new Panel
            {
                Location = new Point(430, y + 15), // Más a la derecha
                Size = new Size(70, 25),
                BackColor = Color.Gray
            };
            this.Controls.Add(pnlStatusIndicator);

            lblStatus = new Label
            {
                Text = "INIT",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };
            pnlStatusIndicator.Controls.Add(lblStatus);

            var btnDelete = new Button
            {
                Text = "ELIMINAR",
                Font = new Font("Segoe UI", 7F, FontStyle.Bold),
                Location = new Point(430, y + 45),
                Size = new Size(70, 25),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 50, 50),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnDelete.FlatAppearance.BorderSize = 0;
            btnDelete.Click += BtnDelete_Click;
            this.Controls.Add(btnDelete);

            y += 80; // Mayor espacio para el Header con Logo muy grande

            // ===== MULTI-PROJECT SELECTOR =====
            var grpProjects = new GroupBox
            {
                Text = "Proyecto Activo",
                Location = new Point(15, y),
                Size = new Size(475, 55)
            };
            this.Controls.Add(grpProjects);

            cboProjects = new ComboBox
            {
                Location = new Point(15, 22),
                Size = new Size(370, 23),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F)
            };
            cboProjects.SelectedIndexChanged += CboProjects_SelectedIndexChanged;
            grpProjects.Controls.Add(cboProjects);

            btnRefreshProjects = new Button
            {
                Location = new Point(395, 20),
                Size = new Size(35, 27),
                FlatStyle = FlatStyle.Flat,
                Image = SvgRefreshIcon(20),
                ImageAlign = ContentAlignment.MiddleCenter
            };
            btnRefreshProjects.FlatAppearance.BorderSize = 1;
            btnRefreshProjects.Click += BtnRefreshProjects_Click;
            grpProjects.Controls.Add(btnRefreshProjects);

            // Indicador de estado de sincronión del proyecto (círculo rojo/verde SVG)
            picProjectSyncStatus = new PictureBox
            {
                Location = new Point(440, 20),
                Size = new Size(24, 24),
                Image = SvgStatusIcon(false, 24),
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            grpProjects.Controls.Add(picProjectSyncStatus);

            y += 65;

            // ===== PATH SELECTION =====
            grpPath = new GroupBox
            {
                Text = "Carpeta de Exportación",
                Location = new Point(15, y),
                Size = new Size(475, 60)
            };
            this.Controls.Add(grpPath);

            txtPath = new TextBox
            {
                Text = _externalPath,
                Location = new Point(15, 25),
                Size = new Size(360, 23),
                ReadOnly = true,
                BackColor = Color.White
            };
            grpPath.Controls.Add(txtPath);

            btnBrowse = new Button
            {
                Text = "...",
                Location = new Point(385, 24),
                Size = new Size(75, 25),
                FlatStyle = FlatStyle.Flat
            };
            btnBrowse.Click += BtnBrowse_Click;
            grpPath.Controls.Add(btnBrowse);

            y += 75;

            // ===== PRECHECK =====
            grpPrecheck = new GroupBox
            {
                Text = "Análisis Previo",
                Location = new Point(15, y),
                Size = new Size(475, 80)
            };
            this.Controls.Add(grpPrecheck);

            int py = 25;
            lblInternal = CreateLabel("Módulos internos:", 15, py, 130, grpPrecheck);
            lblInternalVal = CreateLabel("-", 150, py, 50, grpPrecheck);
            lblExternal = CreateLabel("Archivos externos:", 220, py, 130, grpPrecheck);
            lblExternalVal = CreateLabel("-", 350, py, 50, grpPrecheck);

            py += 25;
            lblExport = CreateLabel("A exportar:", 15, py, 130, grpPrecheck);
            lblExportVal = CreateLabel("-", 150, py, 80, grpPrecheck);
            lblImport = CreateLabel("A importar:", 220, py, 130, grpPrecheck);
            lblImportVal = CreateLabel("-", 350, py, 80, grpPrecheck);

            y += 95; // Move below Análisis Previo

            // ===== MANUAL BUTTONS =====
            var btnExportAll = new Button
            {
                Text = "EXPORTAR TODO",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(15, y),
                Size = new Size(225, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 150, 80),
                ForeColor = Color.White,
                Image = SvgExportIcon(24),
                ImageAlign = ContentAlignment.MiddleCenter,
                TextAlign = ContentAlignment.MiddleCenter,
                TextImageRelation = TextImageRelation.ImageBeforeText
            };
            btnExportAll.FlatAppearance.BorderSize = 0;
            btnExportAll.Click += BtnExportAll_Click;
            this.Controls.Add(btnExportAll);

            var btnImportAll = new Button
            {
                Text = "IMPORTAR TODO",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(250, y),
                Size = new Size(240, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(150, 100, 0),
                ForeColor = Color.White,
                Image = SvgImportIcon(24),
                ImageAlign = ContentAlignment.MiddleCenter,
                TextAlign = ContentAlignment.MiddleCenter,
                TextImageRelation = TextImageRelation.ImageBeforeText
            };
            btnImportAll.FlatAppearance.BorderSize = 0;
            btnImportAll.Click += BtnImportAll_Click;
            this.Controls.Add(btnImportAll);

            y += 45;

            // Auto-Sync Toggle - Functional ON/OFF button
            btnAutoSync = new Button
            {
                Text = " AUTO-SYNC: OFF",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(15, y),
                Size = new Size(350, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(100, 100, 100),
                ForeColor = Color.White
            };
            btnAutoSync.FlatAppearance.BorderSize = 0;
            btnAutoSync.Click += BtnAutoSync_Click;
            this.Controls.Add(btnAutoSync);

            // Stop Button
            var btnStop = new Button
            {
                Text = " DETENER",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(375, y),
                Size = new Size(115, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 50, 50),
                ForeColor = Color.White,
                Image = SvgStopIcon(24),
                ImageAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText
            };
            btnStop.FlatAppearance.BorderSize = 0;
            btnStop.Click += BtnStop_Click;
            this.Controls.Add(btnStop);

            y += 45;

            // ===== MESSAGE =====
            lblMessage = new Label
            {
                Text = "",
                Location = new Point(15, y),
                Size = new Size(320, 20),
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            this.Controls.Add(lblMessage);

            lblElapsed = new Label
            {
                Text = "Tiempo: --",
                Location = new Point(350, y),
                Size = new Size(140, 20),
                ForeColor = Color.FromArgb(100, 100, 100),
                TextAlign = ContentAlignment.MiddleRight
            };
            this.Controls.Add(lblElapsed);

            y += 25;

            // ===== CERTIFICATE =====
            grpCertificate = new GroupBox
            {
                Text = "Certificado de Integridad",
                Location = new Point(15, y),
                Size = new Size(475, 80),
                Visible = false
            };
            this.Controls.Add(grpCertificate);

            py = 20;
            CreateLabel("Estado:", 15, py, 70, grpCertificate);
            lblCertStatusVal = CreateLabel("-", 90, py, 100, grpCertificate);
            CreateLabel("Integridad:", 200, py, 80, grpCertificate);
            lblCertIntegrityVal = CreateLabel("-", 280, py, 100, grpCertificate);

            py += 22;
            CreateLabel("Limpieza:", 15, py, 70, grpCertificate);
            lblCertCleanupVal = CreateLabel("-", 90, py, 100, grpCertificate);
            CreateLabel("Hash:", 200, py, 80, grpCertificate);
            lblCertHashVal = CreateLabel("-", 280, py, 180, grpCertificate);
            lblCertHashVal.Font = new Font("Consolas", 8F);

            y += 10; // Reducido espacio de grupo oculto (era 95)

            // ===== TABS: LOGS & HISTORY =====
            tabLogs = new TabControl
            {
                Location = new Point(15, y),
                Size = new Size(475, 250)
            };
            this.Controls.Add(tabLogs);

            // Tab: Logs
            var tabLog = new TabPage("Logs");
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White,
                Font = new Font("Consolas", 8F)
            };
            tabLog.Controls.Add(txtLog);
            tabLogs.TabPages.Add(tabLog);

            // Tab: Change History con mejoras visuales
            var tabHistory = new TabPage("Historial de Cambios");
            lstChanges = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("Segoe UI", 9F),
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            lstChanges.Columns.Add("", 30);           // Icono
            lstChanges.Columns.Add("Hora", 70);
            lstChanges.Columns.Add("Acción", 85);
            lstChanges.Columns.Add("Módulo", 140);
            lstChanges.Columns.Add("Resultado", 120);
            tabHistory.Controls.Add(lstChanges);
            tabLogs.TabPages.Add(tabHistory);

            y += 260; // Altura TabControl (250) + Margen (10)

            // ===== CLOSE BUTTON =====
            btnClose = new Button
            {
                Text = "Cerrar",
                Location = new Point(210, y),
                Size = new Size(90, 30),
                FlatStyle = FlatStyle.Flat
            };
            btnClose.Click += (s, e) => this.Hide();
            this.Controls.Add(btnClose);

            // ===== IMAGE LIST FOR HISTORY =====
            imgActions = new ImageList { ImageSize = new Size(16, 16), ColorDepth = ColorDepth.Depth32Bit };
            try {
                imgActions.Images.Add("EXPORT", SvgExportIcon(16));
                imgActions.Images.Add("IMPORT", SvgImportIcon(16));
                imgActions.Images.Add("AUTO", SvgAutoIcon(16));
                imgActions.Images.Add("ERROR", SvgErrorIcon(16));
                imgActions.Images.Add("STOP", SvgStopIcon(16));
                imgActions.Images.Add("PLAY", SvgPlayIcon(16));
                imgActions.Images.Add("DELETE", SvgDeleteIcon(16));
                imgActions.Images.Add("LOG", SvgLogIcon(16));
            } catch { }
            
            lstChanges.SmallImageList = imgActions;
        }

        private Label CreateLabel(string text, int x, int y, int width, Control parent)
        {
            var lbl = new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, 20),
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            parent.Controls.Add(lbl);
            return lbl;
        }

        private Bitmap SvgExportIcon(int size = 24)
        {
            // Icono de Exportación: Flecha arriba con degradado esmeralda vibrante
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                
                float padding = size * 0.1f;
                RectangleF rect = new RectangleF(padding, padding, size - 2 * padding, size - 2 * padding);
                
                // Brillo de fondo (Glow)
                using (var glowBrush = new System.Drawing.Drawing2D.PathGradientBrush(new PointF[] { 
                    new PointF(size/2f, size/2f), new PointF(0,0), new PointF(size, 0), new PointF(size, size), new PointF(0, size) 
                }))
                {
                    glowBrush.CenterColor = Color.FromArgb(40, 0, 255, 127);
                    glowBrush.SurroundColors = new Color[] { Color.Transparent };
                }

                // Flecha (Cuerpo y Triángulo)
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                float w = rect.Width;
                float h = rect.Height;
                float cx = size / 2f;
                
                // Puntos de la flecha
                PointF[] points = {
                    new PointF(cx, rect.Top),            // Punta
                    new PointF(rect.Right, size * 0.55f), // Derecha ala
                    new PointF(cx + w*0.2f, size * 0.55f),// Hombro derecho
                    new PointF(cx + w*0.2f, rect.Bottom), // Base derecha
                    new PointF(cx - w*0.2f, rect.Bottom), // Base izquierda
                    new PointF(cx - w*0.2f, size * 0.55f),// Hombro izquierdo
                    new PointF(rect.Left, size * 0.55f)   // Izquierda ala
                };
                path.AddPolygon(points);

                using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(rect, 
                    Color.FromArgb(0, 255, 127), Color.FromArgb(0, 150, 80), 45f))
                {
                    g.FillPath(brush, path);
                }

                using (var pen = new Pen(Color.FromArgb(200, 255, 255, 255), 1.5f))
                {
                    g.DrawPath(pen, path);
                }
            }
            return bitmap;
        }

        private Bitmap SvgImportIcon(int size = 24)
        {
            // Icono de Importación: Flecha abajo con degradado naranja/oro vibrante
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                
                float padding = size * 0.1f;
                RectangleF rect = new RectangleF(padding, padding, size - 2 * padding, size - 2 * padding);
                
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                float w = rect.Width;
                float h = rect.Height;
                float cx = size / 2f;
                
                // Puntos de la flecha invertida
                PointF[] points = {
                    new PointF(cx, rect.Bottom),         // Punta abajo
                    new PointF(rect.Left, size * 0.45f), // Izquierda ala
                    new PointF(cx - w*0.2f, size * 0.45f),// Hombro izquierdo
                    new PointF(cx - w*0.2f, rect.Top),    // Base superior izquierda
                    new PointF(cx + w*0.2f, rect.Top),    // Base superior derecha
                    new PointF(cx + w*0.2f, size * 0.45f),// Hombro derecho
                    new PointF(rect.Right, size * 0.45f) // Derecha ala
                };
                path.AddPolygon(points);

                using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(rect, 
                    Color.FromArgb(255, 180, 0), Color.FromArgb(255, 100, 0), 45f))
                {
                    g.FillPath(brush, path);
                }

                using (var pen = new Pen(Color.FromArgb(200, 255, 255, 255), 1.5f))
                {
                    g.DrawPath(pen, path);
                }
            }
            return bitmap;
        }

        private Bitmap SvgRefreshIcon(int size = 20)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                float cx = size / 2f;
                float cy = size / 2f;
                float radius = size * 0.38f;
                float thickness = size * 0.14f;

                // Color degradado azul vibrante
                Color colorMain = Color.FromArgb(0, 150, 220);
                Color colorDark = Color.FromArgb(0, 100, 180);

                // Arco superior (flecha derecha hacia arriba)
                using (var pen = new Pen(colorMain, thickness))
                {
                    pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    g.DrawArc(pen, cx - radius, cy - radius, radius * 2, radius * 2, -45, 180);
                }

                // Arco inferior (flecha izquierda hacia abajo)
                using (var pen = new Pen(colorDark, thickness))
                {
                    pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    g.DrawArc(pen, cx - radius, cy - radius, radius * 2, radius * 2, 135, 180);
                }

                // Flecha superior (punta hacia arriba-derecha)
                float arrowSize = size * 0.22f;
                using (var brush = new SolidBrush(colorMain))
                {
                    float ax = cx + radius * 0.7f;
                    float ay = cy - radius * 0.7f;
                    PointF[] arrow1 = {
                        new PointF(ax, ay - arrowSize * 0.8f),
                        new PointF(ax + arrowSize * 0.7f, ay + arrowSize * 0.3f),
                        new PointF(ax - arrowSize * 0.4f, ay + arrowSize * 0.3f)
                    };
                    g.FillPolygon(brush, arrow1);
                }

                // Flecha inferior (punta hacia abajo-izquierda)
                using (var brush = new SolidBrush(colorDark))
                {
                    float ax = cx - radius * 0.7f;
                    float ay = cy + radius * 0.7f;
                    PointF[] arrow2 = {
                        new PointF(ax, ay + arrowSize * 0.8f),
                        new PointF(ax - arrowSize * 0.7f, ay - arrowSize * 0.3f),
                        new PointF(ax + arrowSize * 0.4f, ay - arrowSize * 0.3f)
                    };
                    g.FillPolygon(brush, arrow2);
                }
            }
            return bitmap;
        }

        private Bitmap SvgStatusIcon(bool isActive, int size = 20)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                float margin = size * 0.1f;
                RectangleF rect = new RectangleF(margin, margin, size - 2 * margin, size - 2 * margin);

                // Colores según estado
                Color colorLight, colorDark;
                if (isActive)
                {
                    colorLight = Color.FromArgb(100, 220, 100);  // Verde claro
                    colorDark = Color.FromArgb(0, 150, 50);       // Verde oscuro
                }
                else
                {
                    colorLight = Color.FromArgb(255, 100, 100);  // Rojo claro
                    colorDark = Color.FromArgb(180, 30, 30);      // Rojo oscuro
                }

                // Círculo con degradado radial
                using (var path = new System.Drawing.Drawing2D.GraphicsPath())
                {
                    path.AddEllipse(rect);
                    
                    using (var brush = new System.Drawing.Drawing2D.PathGradientBrush(path))
                    {
                        brush.CenterColor = colorLight;
                        brush.SurroundColors = new[] { colorDark };
                        brush.CenterPoint = new PointF(size * 0.35f, size * 0.35f);
                        g.FillPath(brush, path);
                    }

                    // Borde sutil
                    using (var pen = new Pen(Color.FromArgb(100, 0, 0, 0), 1f))
                    {
                        g.DrawEllipse(pen, rect);
                    }
                }

                // Brillo superior (efecto 3D)
                RectangleF glowRect = new RectangleF(
                    margin + size * 0.15f, 
                    margin + size * 0.1f, 
                    size * 0.4f, 
                    size * 0.25f);
                using (var glowBrush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    glowRect, 
                    Color.FromArgb(150, 255, 255, 255), 
                    Color.FromArgb(0, 255, 255, 255), 
                    90f))
                {
                    g.FillEllipse(glowBrush, glowRect);
                }
            }
            return bitmap;
        }

        private Bitmap SvgDeleteIcon(int size = 24)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                
                float padding = size * 0.15f;
                RectangleF rect = new RectangleF(padding, padding, size - 2 * padding, size - 2 * padding);
                
                Color trashColor = Color.FromArgb(220, 50, 50);
                
                using (var pen = new Pen(trashColor, size * 0.08f))
                {
                    pen.LineJoin = System.Drawing.Drawing2D.LineJoin.Round;
                    pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;

                    // Tapa
                    g.DrawLine(pen, rect.Left + rect.Width * 0.1f, rect.Top + rect.Height * 0.15f, rect.Right - rect.Width * 0.1f, rect.Top + rect.Height * 0.15f);
                    // Asa de la tapa
                    g.DrawArc(pen, rect.Left + rect.Width * 0.35f, rect.Top, rect.Width * 0.3f, rect.Height * 0.2f, 180, 180);
                    
                    // Cuerpo
                    PointF[] body = {
                        new PointF(rect.Left + rect.Width * 0.2f, rect.Top + rect.Height * 0.2f),
                        new PointF(rect.Right - rect.Width * 0.2f, rect.Top + rect.Height * 0.2f),
                        new PointF(rect.Right - rect.Width * 0.25f, rect.Bottom),
                        new PointF(rect.Left + rect.Width * 0.25f, rect.Bottom)
                    };
                    g.FillPolygon(new SolidBrush(Color.FromArgb(40, trashColor)), body);
                    g.DrawPolygon(pen, body);
                    
                    // Líneas verticales
                    g.DrawLine(pen, rect.Left + rect.Width * 0.4f, rect.Top + rect.Height * 0.4f, rect.Left + rect.Width * 0.42f, rect.Bottom - rect.Height * 0.15f);
                    g.DrawLine(pen, rect.Left + rect.Width * 0.6f, rect.Top + rect.Height * 0.4f, rect.Left + rect.Width * 0.58f, rect.Bottom - rect.Height * 0.15f);
                }
            }
            return bitmap;
        }

        private Bitmap SvgStopIcon(int size = 24)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                float p = size * 0.2f;
                using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(new RectangleF(p, p, size - 2 * p, size - 2 * p), 
                    Color.FromArgb(255, 80, 80), Color.FromArgb(180, 0, 0), 45f))
                {
                    g.FillRectangle(brush, p, p, size - 2 * p, size - 2 * p);
                }
                using (var pen = new Pen(Color.FromArgb(150, 255, 255, 255), 1.5f))
                {
                    g.DrawRectangle(pen, p, p, size - 2 * p, size - 2 * p);
                }
            }
            return bitmap;
        }

        private Bitmap SvgPlayIcon(int size = 24)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                float p = size * 0.2f;
                PointF[] pts = { new PointF(p, p), new PointF(size - p, size / 2f), new PointF(p, size - p) };
                using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(new RectangleF(p, p, size - 2 * p, size - 2 * p), 
                    Color.FromArgb(80, 200, 80), Color.FromArgb(0, 150, 0), 45f))
                {
                    g.FillPolygon(brush, pts);
                }
                using (var pen = new Pen(Color.FromArgb(150, 255, 255, 255), 1.5f))
                {
                    g.DrawPolygon(pen, pts);
                }
            }
            return bitmap;
        }

        private Bitmap SvgAutoIcon(int size = 24)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                float p = size * 0.15f;
                PointF[] pts = { 
                    new PointF(size * 0.6f, p), 
                    new PointF(size * 0.2f, size * 0.55f), 
                    new PointF(size * 0.45f, size * 0.55f), 
                    new PointF(size * 0.4f, size - p), 
                    new PointF(size * 0.8f, size * 0.45f), 
                    new PointF(size * 0.55f, size * 0.45f) 
                };
                using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(new RectangleF(p, p, size - 2 * p, size - 2 * p), 
                    Color.FromArgb(255, 255, 0), Color.FromArgb(255, 150, 0), 45f))
                {
                    g.FillPolygon(brush, pts);
                }
                using (var pen = new Pen(Color.FromArgb(100, 0, 0, 0), 1f))
                {
                    g.DrawPolygon(pen, pts);
                }
            }
            return bitmap;
        }

        private Bitmap SvgErrorIcon(int size = 24)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                float p = size * 0.15f;
                RectangleF rect = new RectangleF(p, p, size - 2 * p, size - 2 * p);
                using (var pen = new Pen(Color.FromArgb(200, 50, 50), size * 0.15f))
                {
                    pen.StartCap = pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    g.DrawLine(pen, rect.Left, rect.Top, rect.Right, rect.Bottom);
                    g.DrawLine(pen, rect.Right, rect.Top, rect.Left, rect.Bottom);
                }
            }
            return bitmap;
        }

        private Bitmap SvgLogIcon(int size = 24)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                float p = size * 0.15f;
                using (var pen = new Pen(Color.FromArgb(100, 100, 100), size * 0.08f))
                {
                    g.DrawRectangle(pen, p, p, size - 2 * p, size - 2 * p);
                    g.DrawLine(pen, p * 2.5f, p * 3f, size - p * 2.5f, p * 3f);
                    g.DrawLine(pen, p * 2.5f, p * 5f, size - p * 2.5f, p * 5f);
                    g.DrawLine(pen, p * 2.5f, p * 7f, size - p * 4f, p * 7f);
                }
            }
            return bitmap;
        }


        public void PublicRefreshUI(string moduleName, string action)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action<string, string>(PublicRefreshUI), moduleName, action);
                return;
            }

            AppendLog($"AUTO: {action} - {moduleName}");
            AddChangeRecord("AUTO", moduleName, action);
            
            // Refresh counts
            int intCount = 0;
            if (_vbaProject != null)
            {
                foreach (VBIDE.VBComponent comp in _vbaProject.VBComponents)
                {
                    if (comp.Type != VBIDE.vbext_ComponentType.vbext_ct_Document)
                        intCount++;
                }
            }

            int extCount = 0;
            if (Directory.Exists(_externalPath))
            {
                extCount = Directory.GetFiles(_externalPath, "*.bas", SearchOption.AllDirectories).Length +
                          Directory.GetFiles(_externalPath, "*.cls", SearchOption.AllDirectories).Length +
                          Directory.GetFiles(_externalPath, "*.frm", SearchOption.AllDirectories).Length;
            }

            lblInternalVal.Text = intCount.ToString();
            lblExternalVal.Text = extCount.ToString();
        }

        #endregion

        #region Form Events

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            ValidateEnvironment();
            RefreshProjectList(); // Cargar lista de proyectos al inicio
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Ocultar el formulario en lugar de cerrarlo completamente
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; // Cancelar el cierre
                this.Hide(); // Ocultar el formulario
            }
            else
            {
                CleanupEngine();
                base.OnFormClosing(e);
            }
        }

        #endregion

        #region Browse Folder (Modern Dialog)

        // COM interfaces for modern folder picker (Vista+)
        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        private class FileOpenDialogInternal { }

        [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileDialog
        {
            [PreserveSig] int Show([In] IntPtr parent);
            void SetFileTypes();  // cycled cycled cycled
            void SetFileTypeIndex([In] uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(); // cycled cycled cycled
            void Unadvise([In] uint dwCookie);
            void SetOptions([In] uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            void SetFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi, int alignment);
            void SetDefaultExtension([In, MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close([MarshalAs(UnmanagedType.Error)] int hr);
            void SetClientGuid([In] ref Guid guid);
            void ClearClientData();
            void SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(); // cycled cycled cycled
            void GetParent(); // cycled cycled cycled
            void GetDisplayName([In] uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
            void GetAttributes();  // cycled cycled cycled
            void Compare();  // cycled cycled cycled
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            [In] IntPtr pbc,
            [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [Out, MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

        private const uint FOS_PICKFOLDERS = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint FOS_NOVALIDATE = 0x00000100;
        private const uint SIGDN_FILESYSPATH = 0x80058000;

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            string? selectedPath = ShowModernFolderDialog(_externalPath);
            
            if (!string.IsNullOrEmpty(selectedPath))
            {
                _externalPath = selectedPath!;
                txtPath.Text = _externalPath;
                
                // Update path in AddInHost for auto-sync
                if (_host != null)
                    _host.SetExternalPath(_externalPath);
                
                AppendLog($"Carpeta cambiada: {_externalPath}");
                SetupUIWatcher(); // Reiniciar watcher con la nueva ruta
                ExecutePrecheck();
            }
        }

        private string? ShowModernFolderDialog(string initialPath)
        {
            try
            {
                // Try modern dialog (Vista+)
                var dialog = (IFileDialog)new FileOpenDialogInternal();
                dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_NOVALIDATE);
                dialog.SetTitle("Seleccione la carpeta de sincronización");

                // Set initial folder
                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath))
                {
                    IShellItem folder;
                    SHCreateItemFromParsingName(initialPath, IntPtr.Zero, typeof(IShellItem).GUID, out folder);
                    dialog.SetFolder(folder);
                }

                // Show dialog
                int hr = dialog.Show(this.Handle);
                if (hr == 0) // S_OK
                {
                    dialog.GetResult(out IShellItem result);
                    result.GetDisplayName(SIGDN_FILESYSPATH, out string path);
                    return path;
                }
            }
            catch
            {
                // Fallback to old dialog if modern fails
                using (var fallback = new FolderBrowserDialog())
                {
                    fallback.Description = "Seleccione la carpeta de exportación";
                    fallback.SelectedPath = initialPath;
                    if (fallback.ShowDialog() == DialogResult.OK)
                        return fallback.SelectedPath;
                }
            }

            return null;
        }

        #endregion

        #region Validation

        private void ValidateEnvironment()
        {
            AppendLog("Validando entorno...");

            try
            {
                if (_vbaProject != null)
                {
                    int count = _vbaProject.VBComponents.Count;
                    AppendLog($"Proyecto VBA accesible. Componentes: {count}");
                    SetState(UIState.Validated);
                    AppendLog("Entorno validado correctamente");
                }
                else
                {
                    AppendLog("AVISO: No se detectó proyecto activo. Use 'Actualizar Proyectos' si el archivo es nuevo.");
                    SetState(UIState.Init); // Use Init instead of Ready
                }

                if (!Directory.Exists(_externalPath))
                {
                    Directory.CreateDirectory(_externalPath);
                    AppendLog($"Carpeta creada: {_externalPath}");
                }

                ExecutePrecheck();
            }
            catch (Exception ex)
            {
                SetState(UIState.RolledBack);
                AppendLog($"ERROR de acceso a VBA: {ex.Message}");
                lblMessage.Text = "¿Confianza al modelo de objetos de VBA habilitada?";
            }
        }

        private void ExecutePrecheck()
        {
            AppendLog("Ejecutando precheck...");

            int intCount = 0;
            if (_vbaProject != null)
            {
                foreach (VBIDE.VBComponent comp in _vbaProject.VBComponents)
                {
                    intCount++;
                }
            }
            else
            {
                AppendLog("AVISO: No se detectó un proyecto VBA activo.");
            }

            int extCount = 0;
            if (Directory.Exists(_externalPath))
            {
                extCount = Directory.GetFiles(_externalPath, "*.bas", SearchOption.AllDirectories).Length +
                          Directory.GetFiles(_externalPath, "*.cls", SearchOption.AllDirectories).Length +
                          Directory.GetFiles(_externalPath, "*.frm", SearchOption.AllDirectories).Length;
            }

            lblInternalVal.Text = intCount.ToString();
            lblExternalVal.Text = extCount.ToString();
            lblExportVal.Text = intCount + " posibles";
            lblImportVal.Text = extCount + " posibles";

            SetState(UIState.Analyzed);
            EnableButtons();
            AppendLog("Precheck completado");
        }

        // Se han eliminado bloques obsoletos.

        #endregion

        // Se han eliminado bloques obsoletos.

        #region Auto-Sync

        private void BtnAutoSync_Click(object? sender, EventArgs e)
        {
            if (_host == null)
            {
                AppendLog("Host no disponible para auto-sync");
                return;
            }

            // Obtener el proyecto actualmente seleccionado
            VBIDE.VBProject? selectedProject = _vbaProject;
            if (cboProjects.SelectedItem is ProjectDisplayItem item && item.Project != null)
            {
                selectedProject = item.Project;
            }

            if (selectedProject == null)
            {
                AppendLog("ERROR: No hay proyecto seleccionado");
                return;
            }

            // Verificar si este proyecto específico tiene sincronización activa
            bool isProjectSyncing = _host.IsProjectSyncing(selectedProject);

            if (isProjectSyncing)
            {
                // Turn OFF - Detener sincronización para este proyecto
                _host.StopSyncForProject(selectedProject);
                btnAutoSync.Text = " AUTO-SYNC: APAGADO";
                btnAutoSync.BackColor = Color.FromArgb(180, 50, 50);
                try { picProjectSyncStatus.Image = SvgStatusIcon(false, 24); } catch { }
                AppendLog($"Auto-Sync DESACTIVADO para: {selectedProject.Name}");
                AddChangeRecord("AUTOSYNC", selectedProject.Name, "Detenido");
            }
            else
            {
                // Turn ON - Iniciar sincronización para este proyecto
                
                // Sync path to host FIRST (para compatibilidad)
                _host.SetExternalPath(_externalPath);
                
                AppendLog($"Exportando módulos de {selectedProject.Name}...");
                
                // Export all modules first
                int count = 0;
                try
                {
                    foreach (VBIDE.VBComponent comp in selectedProject.VBComponents)
                    {
                        string ext = GetExtension(comp.Type);
                        string subfolder = GetSubfolder(comp.Type);
                        string targetDir = Path.Combine(_externalPath, subfolder);
                        
                        if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
                        
                        string path = Path.Combine(targetDir, comp.Name + ext);
                        
                        try
                        {
                            string code = "";
                            if (comp.CodeModule.CountOfLines > 0)
                                code = comp.CodeModule.Lines[1, comp.CodeModule.CountOfLines];
                            
                            File.WriteAllText(path, code, new UTF8Encoding(true));
                            count++;
                        }
                        catch { }
                    }
                    AppendLog($"EXPORTADOS: {count} módulos organizada por estructura VBA.");
                }
                catch (Exception ex)
                {
                    AppendLog($"Error exportando: {ex.Message}");
                }
                
                // Iniciar sincronización multi-proyecto
                _host.StartSyncForProject(selectedProject, _externalPath);
                
                // También iniciar el sistema global para compatibilidad
                if (!_host.IsSyncEnabled)
                    _host.StartBackgroundSync();
                
                btnAutoSync.Text = " AUTO-SYNC: ACTIVO";
                btnAutoSync.BackColor = Color.FromArgb(0, 180, 0);
                try { picProjectSyncStatus.Image = SvgStatusIcon(true, 24); } catch { }
                AppendLog($"Auto-Sync ACTIVADO para: {selectedProject.Name} en {_externalPath}");
                AddChangeRecord("AUTOSYNC", selectedProject.Name, $"Iniciado ({count} módulos exportados)");
            }
            ExecutePrecheck(); // Refrescar análisis tras cambio de estado
        }


        private void BtnExportAll_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_vbaProject == null) { AppendLog("ERROR: No hay proyecto VBA"); return; }

                // Obtener la lista de componentes disponibles para exportar
                var components = new List<VBIDE.VBComponent>();
                foreach (VBIDE.VBComponent comp in _vbaProject.VBComponents)
                {
                    components.Add(comp);
                }

                if (components.Count == 0)
                {
                    MessageBox.Show("No hay componentes disponibles para exportar en este proyecto.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Mostrar diálogo para seleccionar componentes a exportar
                using (var form = new SelectComponentsForm(components, "Seleccionar Componentes para Exportar"))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        var selectedComponents = form.SelectedComponents;
                        int count = 0;

                        foreach (VBIDE.VBComponent comp in selectedComponents)
                        {
                            string ext = GetExtension(comp.Type);
                            string subfolder = GetSubfolder(comp.Type);
                            string targetDir = Path.Combine(_externalPath, subfolder);

                            if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

                            string path = Path.Combine(targetDir, comp.Name + ext);

                            try
                            {
                                string code = "";
                                if (comp.CodeModule.CountOfLines > 0)
                                    code = comp.CodeModule.Lines[1, comp.CodeModule.CountOfLines];

                                File.WriteAllText(path, code, new UTF8Encoding(true)); // UTF-8 with BOM
                                count++;
                            }
                            catch { }
                        }

                        AppendLog($"EXPORTADO: {count} módulos organizada por estructura VBA.");
                        AddChangeRecord("EXPORT", "MANUAL", $"{count} módulos");
                        MessageBox.Show($"Exportados {count} módulos a:\n{_externalPath}", "Exportación Completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ExecutePrecheck(); // Refrescar análisis
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("ERROR Export: " + ex.Message);
            }
        }

        private void BtnDelete_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_vbaProject == null) { AppendLog("ERROR: No hay proyecto VBA"); return; }

                var components = new List<VBIDE.VBComponent>();
                foreach (VBIDE.VBComponent comp in _vbaProject.VBComponents)
                {
                    components.Add(comp);
                }

                if (components.Count == 0)
                {
                    MessageBox.Show("No hay componentes para eliminar.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var form = new SelectComponentsForm(components, "Seleccionar Componentes para ELIMINAR"))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        var selected = form.SelectedComponents;
                        if (selected.Count == 0) return;

                        var confirm = MessageBox.Show(
                            "¿Deseas ELIMINAR permanentemente los componentes seleccionados?\n\nSu recuperación es irreversible.",
                            "Confirmar Eliminación",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Warning);

                        if (confirm != DialogResult.OK) return;

                        var doBackup = MessageBox.Show(
                            "¿Deseas crear un respaldo (BACKUP) de seguridad antes de proceder?",
                            "Opción de Respaldo",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        bool shouldBackup = (doBackup == DialogResult.Yes);
                        string? backupPath = null;

                        if (shouldBackup)
                        {
                            backupPath = ShowModernFolderDialog("");
                            if (string.IsNullOrEmpty(backupPath))
                            {
                                AppendLog("Operación CANCELADA: Se requería respaldo pero no se seleccionó carpeta.");
                                return;
                            }
                        }
                        else
                        {
                            var finalWarn = MessageBox.Show(
                                "VAS A ELIMINAR SIN RESPALDO.\n\nEsta acción no se puede deshacer y no habrá copia de seguridad.\n\n¿Estás absolutamente seguro?",
                                "¡ADVERTENCIA CRÍTICA!",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Error);
                            if (finalWarn != DialogResult.Yes) return;
                        }

                        AppendLog($"Iniciando eliminación masiva de {selected.Count} componentes" + (shouldBackup ? " con respaldo..." : " SIN RESPALDO..."));
                        int deletedCount = 0;
                        int backupCount = 0;

                        foreach (VBIDE.VBComponent comp in selected)
                        {
                            string name = comp.Name;
                            var type = comp.Type; // Cache Type
                            string ext = GetExtension(type);
                            string subfolder = GetSubfolder(type);

                            // 1. BACKUP (Opcional)
                            if (shouldBackup && !string.IsNullOrEmpty(backupPath))
                            {
                                try
                                {
                                    string backupFile = Path.Combine(backupPath, subfolder, name + ext);
                                    string? dir = Path.GetDirectoryName(backupFile);
                                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                                    string code = "";
                                    if (comp.CodeModule.CountOfLines > 0)
                                        code = comp.CodeModule.Lines[1, comp.CodeModule.CountOfLines];

                                    File.WriteAllText(backupFile, code, new UTF8Encoding(true));
                                    backupCount++;
                                }
                                catch (Exception ex)
                                {
                                    AppendLog($"ERROR Backup {name}: {ex.Message}");
                                    continue; // Skip deletion if backup requested but failed
                                }
                            }

                            // 2. DELETE INTERNAL (VBA)
                            try
                            {
                                if (type == VBIDE.vbext_ComponentType.vbext_ct_Document)
                                {
                                    if (comp.CodeModule.CountOfLines > 0)
                                        comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines);
                                    AppendLog($"VBA: Código limpiado en {name}");
                                }
                                else
                                {
                                    _vbaProject.VBComponents.Remove(comp);
                                    AppendLog($"VBA: Eliminado {name}");
                                }
                            }
                            catch (Exception ex)
                            {
                                AppendLog($"ERROR VBA Delete {name}: {ex.Message}");
                                continue;
                            }

                            // 3. DELETE EXTERNAL (Disk)
                            try
                            {
                                string extPath = Path.Combine(_externalPath, subfolder, name + ext);
                                if (File.Exists(extPath))
                                {
                                    File.Delete(extPath);
                                    AppendLog($"Disco: Eliminado {name}{ext}");
                                }
                                
                                // Also check for .frx if it's a form
                                if (type == VBIDE.vbext_ComponentType.vbext_ct_MSForm)
                                {
                                    string frxPath = Path.Combine(_externalPath, subfolder, name + ".frx");
                                    if (File.Exists(frxPath)) File.Delete(frxPath);
                                }
                            }
                            catch (Exception ex)
                            {
                                AppendLog($"ERROR Disco Delete {name}: {ex.Message}");
                            }

                            deletedCount++;
                        }

                        string resultMsg = $"Se eliminaron {deletedCount} componentes.";
                        if (shouldBackup) resultMsg += $"\nSe crearon {backupCount} respaldos.";
                        
                        AppendLog($"FIN: {resultMsg}");
                        AddChangeRecord("DELETE", "MANUAL", $"{deletedCount} mod" + (shouldBackup ? $", {backupCount} bak" : ""));
                        MessageBox.Show(resultMsg, "Proceso Completado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                        ExecutePrecheck(); // Refresh counts
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("ERROR Delete: " + ex.Message);
            }
        }

        private void BtnImportAll_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_vbaProject == null) { AppendLog("ERROR: No hay proyecto VBA"); return; }

                // Mostrar diálogo para seleccionar archivos VBA para importar
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Archivos VBA (*.bas;*.cls;*.frm)|*.bas;*.cls;*.frm|Todos los archivos (*.*)|*.*";
                    openFileDialog.Multiselect = true; // Permitir selección múltiple
                    openFileDialog.Title = "Seleccionar archivos VBA para importar";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        int count = 0;
                        foreach (string file in openFileDialog.FileNames)
                        {
                            string name = Path.GetFileNameWithoutExtension(file);
                            string content = File.ReadAllText(file, Encoding.GetEncoding(1252)); // Windows-1252 para acentos

                            try
                            {
                                VBIDE.VBComponent? comp = null;
                                try { comp = _vbaProject.VBComponents.Item(name); } catch { }

                                if (comp == null)
                                {
                                    var type = Path.GetExtension(file).ToLower() == ".cls" || Path.GetExtension(file).ToLower() == ".frm"
                                        ? VBIDE.vbext_ComponentType.vbext_ct_ClassModule
                                        : VBIDE.vbext_ComponentType.vbext_ct_StdModule;

                                    if (Path.GetExtension(file).ToLower() == ".frm")
                                        type = VBIDE.vbext_ComponentType.vbext_ct_MSForm;

                                    comp = _vbaProject.VBComponents.Add(type);
                                    comp.Name = name;
                                }

                                if (comp.CodeModule.CountOfLines > 0)
                                    comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines);

                                comp.CodeModule.AddFromString(content);
                                count++;
                            }
                            catch { }
                        }

                        AppendLog($"IMPORTADO: {count} módulos desde archivos seleccionados");
                        AddChangeRecord("IMPORT", "MANUAL", $"{count} módulos");
                        MessageBox.Show($"Importados {count} módulos desde los archivos seleccionados.", "Importación Completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ExecutePrecheck(); // Refrescar análisis
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("ERROR Import: " + ex.Message);
            }
        }

        private void BtnStop_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_host != null && _host.IsSyncEnabled)
                {
                    _host.StopBackgroundSync();
                    btnAutoSync.Text = " AUTO-SYNC: OFF";
                    btnAutoSync.BackColor = Color.FromArgb(100, 100, 100);
                }

                SetState(UIState.Analyzed);
                EnableButtons();

                AppendLog("Sincronización DETENIDA por el usuario");
                AddChangeRecord("DETENER", "SISTEMA", "Detenido manualmente");
                lblMessage.Text = "Sincronización detenida";
            }
            catch (Exception ex)
            {
                AppendLog($"Error al detener: {ex.Message}");
            }
        }


        #endregion

        #region State Management

        private enum UIState { Init, Validated, Analyzed, Running, Conflict, Committed, RolledBack }

        private void SetState(UIState state)
        {
            _state = state;
            UpdateStateDisplay();
        }

        private void UpdateStateDisplay()
        {
            switch (_state)
            {
                case UIState.Init:
                    lblStatus.Text = "INIT";
                    pnlStatusIndicator.BackColor = Color.Gray;
                    break;
                case UIState.Validated:
                    lblStatus.Text = "VALIDADO";
                    pnlStatusIndicator.BackColor = Color.Green;
                    break;
                case UIState.Analyzed:
                    lblStatus.Text = "LISTO";
                    pnlStatusIndicator.BackColor = Color.Green;
                    break;
                case UIState.Running:
                    lblStatus.Text = "EJECUTANDO";
                    pnlStatusIndicator.BackColor = Color.Orange;
                    break;
                case UIState.Conflict:
                    lblStatus.Text = "CONFLICTO";
                    pnlStatusIndicator.BackColor = Color.Orange;
                    break;
                case UIState.Committed:
                    lblStatus.Text = "COMPLETADO";
                    pnlStatusIndicator.BackColor = Color.Green;
                    break;
                case UIState.RolledBack:
                    lblStatus.Text = "REVERTIDO";
                    pnlStatusIndicator.BackColor = Color.Red;
                    break;
            }
        }

        private void DisableAllActions()
        {
            btnBrowse.Enabled = false;
        }

        private void EnableButtons()
        {
            {
                btnAutoSync.Enabled = true;
                btnBrowse.Enabled = true;
            }
        }

        #endregion

        #region Logging

        private void AppendLog(string message)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
        }

        private void AddChangeRecord(string action, string module, string result)
        {
            var record = new ChangeRecord
            {
                Time = DateTime.Now,
                Action = action,
                Module = module,
                Result = result
            };
            _changeHistory.Add(record);

            // Determinar icono y color según la acción
            string iconKey = GetActionIcon(action);
            Color rowColor = GetActionColor(action, result);

            var item = new ListViewItem("");
            item.ImageKey = iconKey;
            item.SubItems.Add(record.Time.ToString("HH:mm:ss"));
            item.SubItems.Add(record.Action);
            item.SubItems.Add(record.Module);
            item.SubItems.Add(record.Result);
            item.ForeColor = rowColor;

            // Alternar color de fondo para mejor legibilidad
            if (lstChanges.Items.Count % 2 == 0)
                item.BackColor = Color.FromArgb(248, 248, 248);

            lstChanges.Items.Insert(0, item);

            // Limitar historial visible a 100 entradas para rendimiento
            if (lstChanges.Items.Count > 100)
                lstChanges.Items.RemoveAt(lstChanges.Items.Count - 1);
        }

        private string GetActionIcon(string action)
        {
            return action.ToUpperInvariant() switch
            {
                "EXPORT" => "EXPORT",
                "EXPORTAR" => "EXPORT",
                "EXPORTADO" => "EXPORT",
                "IMPORT" => "IMPORT",
                "IMPORTAR" => "IMPORT",
                "IMPORTADO" => "IMPORT",
                "AUTO" => "AUTO",
                "AUTOSYNC" => "AUTO",
                "ERROR" => "ERROR",
                "DETENER" => "STOP",
                "INICIAR" => "PLAY",
                "DELETE" => "DELETE",
                _ => "LOG"
            };
        }

        private Color GetActionColor(string action, string result)
        {
            if (result.Contains("ERROR") || result.Contains("FATAL"))
                return Color.FromArgb(200, 50, 50);  // Rojo oscuro
            if (result.Contains("ROLLBACK"))
                return Color.FromArgb(200, 120, 0);  // Naranja
            
            return action.ToUpperInvariant() switch
            {
                "EXPORT" => Color.FromArgb(0, 120, 80),
                "EXPORTAR" => Color.FromArgb(0, 120, 80),
                "EXPORTADO" => Color.FromArgb(0, 120, 80),
                "IMPORT" => Color.FromArgb(0, 90, 180),
                "IMPORTAR" => Color.FromArgb(0, 90, 180),
                "IMPORTADO" => Color.FromArgb(0, 90, 180),
                "AUTO" => Color.FromArgb(100, 50, 150),
                "AUTOSYNC" => Color.FromArgb(0, 150, 100),
                "DETENER" => Color.FromArgb(150, 50, 50),
                "DELETE" => Color.FromArgb(180, 50, 50),
                _ => Color.FromArgb(80, 80, 80)
            };
        }



        #endregion

        #region Multi-Project Events

        private void BtnRefreshProjects_Click(object? sender, EventArgs e)
        {
            RefreshProjectList();
        }

        private void RefreshProjectList()
        {
            try
            {
                VBIDE.VBE? vbe = null;
                if (_host != null) vbe = _host.VBE;
                if (vbe == null && _vbaProject != null) vbe = _vbaProject.VBE;

                if (vbe == null)
                {
                    AppendLog("AVISO: No se pudo obtener el motor VBE.");
                    return;
                }

                cboProjects.Items.Clear(); // Limpiar para evitar duplicados al refrescar

                // Enumerar todos los proyectos VBA abiertos
                foreach (VBIDE.VBProject proj in vbe.VBProjects)
                {
                    try
                    {
                        string displayName = $"{proj.Name}";
                        
                        // Intentar obtener el nombre del archivo si está disponible
                        try
                        {
                            if (!string.IsNullOrEmpty(proj.FileName))
                                displayName = $"{proj.Name} ({Path.GetFileName(proj.FileName)})";
                        }
                        catch { }
                        
                        cboProjects.Items.Add(new ProjectDisplayItem { Project = proj, DisplayName = displayName });
                        
                        // Seleccionar el proyecto actual
                        if (_vbaProject != null && proj.Name == _vbaProject.Name)
                            cboProjects.SelectedIndex = cboProjects.Items.Count - 1;
                    }
                    catch { }
                }

                if (cboProjects.Items.Count > 0 && cboProjects.SelectedIndex < 0)
                    cboProjects.SelectedIndex = 0;

                ExecutePrecheck(); // También refrescar el análisis al refrescar proyectos
                AppendLog($"Proyectos detectados: {cboProjects.Items.Count}");
                AddChangeRecord("REFRESH", "PROYECTOS", $"{cboProjects.Items.Count} proyectos");
            }
            catch (Exception ex)
            {
                AppendLog($"Error al refrescar proyectos: {ex.Message}");
            }
        }

        private void CboProjects_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cboProjects.SelectedItem is ProjectDisplayItem item && item.Project != null)
            {
                _vbaProject = item.Project;
                AppendLog($"Proyecto seleccionado: {item.DisplayName}");
                
                // Actualizar contadores para el proyecto seleccionado
                try
                {
                    // Si el host tiene una ruta específica para este proyecto, usarla
                    if (_host != null)
                    {
                        var syncPath = _host.GetProjectSyncPath(item.Project);
                        if (!string.IsNullOrEmpty(syncPath))
                        {
                            _externalPath = syncPath!;
                            txtPath.Text = syncPath;
                        }

                        bool isSyncing = _host.IsProjectSyncing(item.Project);
                        picProjectSyncStatus.Image = SvgStatusIcon(isSyncing, 24);
                    }

                    // Forzar recalculo de todo el análisis previo
                    ExecutePrecheck();
                }
                catch { }
            }
        }


        // Clase auxiliar para mostrar proyectos en el ComboBox
        private class ProjectDisplayItem
        {
            public VBIDE.VBProject? Project { get; set; }
            public string DisplayName { get; set; } = string.Empty;
            
            public override string ToString() => DisplayName;
        }

        #endregion

        #region Cleanup

        private void CleanupEngine()
        {
            // Nothing to clean manually. AddInHost handles lifecycle.
        }

        #endregion

        #region Helpers

        private string GetSubfolder(VBIDE.vbext_ComponentType type)
        {
            switch (type)
            {
                case VBIDE.vbext_ComponentType.vbext_ct_StdModule: return "Módulos";
                case VBIDE.vbext_ComponentType.vbext_ct_MSForm: return "Formularios";
                case VBIDE.vbext_ComponentType.vbext_ct_ClassModule: return "Módulos de clase";
                case VBIDE.vbext_ComponentType.vbext_ct_Document: return "Microsoft Excel Objetos";
                default: return "Otros";
            }
        }

        private static string GetExtension(VBIDE.vbext_ComponentType type)
        {
            return type switch
            {
                VBIDE.vbext_ComponentType.vbext_ct_ClassModule => ".cls",
                VBIDE.vbext_ComponentType.vbext_ct_MSForm => ".frm",
                VBIDE.vbext_ComponentType.vbext_ct_Document => ".cls", // ThisWorkbook, Sheets
                _ => ".bas"
            };
        }

        private static bool IsSupportedExtension(string ext)
        {
            return ext.ToLowerInvariant() switch
            {
                ".bas" => true,
                ".cls" => true,
                ".frm" => true,
                _ => false
            };
        }

        #endregion

        #region Nested Types

        private class ChangeRecord
        {
            public DateTime Time { get; set; }
            public string Action { get; set; } = string.Empty;
            public string Module { get; set; } = string.Empty;
            public string Result { get; set; } = string.Empty;
        }

        #endregion
    }
}
