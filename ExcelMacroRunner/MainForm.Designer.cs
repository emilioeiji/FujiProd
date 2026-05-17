namespace ExcelMacroRunner;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;
    private Button btnStart;
    private Button btnStop;
    private Button btnRestartNow;
    private Button btnKillExcel;
    private Button btnOpenLogFolder;
    private Label lblStatusTitle;
    private Label lblStatusValue;
    private Label lblLastExecutionTitle;
    private Label lblLastExecutionValue;
    private Label lblHtmlLastUpdateTitle;
    private Label lblHtmlLastUpdateValue;
    private Label lblHtmlPathTitle;
    private Label lblHtmlPathValue;
    private Label lblNextActionTitle;
    private Label lblNextActionValue;
    private Label lblCurrentMessageTitle;
    private TextBox txtCurrentMessage;
    private TextBox txtLogs;
    private Label lblLogsTitle;
    private TableLayoutPanel tableMain;
    private FlowLayoutPanel panelButtons;
    private TableLayoutPanel tableStatus;
    private Label lblConfigPathTitle;
    private Label lblConfigPathValue;
    private Label lblLogFolderTitle;
    private Label lblLogFolderValue;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }

        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        btnStart = new Button();
        btnStop = new Button();
        btnRestartNow = new Button();
        btnKillExcel = new Button();
        btnOpenLogFolder = new Button();
        lblStatusTitle = new Label();
        lblStatusValue = new Label();
        lblLastExecutionTitle = new Label();
        lblLastExecutionValue = new Label();
        lblHtmlLastUpdateTitle = new Label();
        lblHtmlLastUpdateValue = new Label();
        lblHtmlPathTitle = new Label();
        lblHtmlPathValue = new Label();
        lblNextActionTitle = new Label();
        lblNextActionValue = new Label();
        lblCurrentMessageTitle = new Label();
        txtCurrentMessage = new TextBox();
        txtLogs = new TextBox();
        lblLogsTitle = new Label();
        tableMain = new TableLayoutPanel();
        panelButtons = new FlowLayoutPanel();
        tableStatus = new TableLayoutPanel();
        lblConfigPathTitle = new Label();
        lblConfigPathValue = new Label();
        lblLogFolderTitle = new Label();
        lblLogFolderValue = new Label();
        tableMain.SuspendLayout();
        panelButtons.SuspendLayout();
        tableStatus.SuspendLayout();
        SuspendLayout();
        // 
        // btnStart
        // 
        btnStart.AutoSize = true;
        btnStart.Location = new Point(3, 3);
        btnStart.Name = "btnStart";
        btnStart.Size = new Size(100, 35);
        btnStart.TabIndex = 0;
        btnStart.Text = "Iniciar";
        btnStart.UseVisualStyleBackColor = true;
        btnStart.Click += btnStart_Click;
        // 
        // btnStop
        // 
        btnStop.AutoSize = true;
        btnStop.Location = new Point(109, 3);
        btnStop.Name = "btnStop";
        btnStop.Size = new Size(100, 35);
        btnStop.TabIndex = 1;
        btnStop.Text = "Parar";
        btnStop.UseVisualStyleBackColor = true;
        btnStop.Click += btnStop_Click;
        // 
        // btnRestartNow
        // 
        btnRestartNow.AutoSize = true;
        btnRestartNow.Location = new Point(215, 3);
        btnRestartNow.Name = "btnRestartNow";
        btnRestartNow.Size = new Size(130, 35);
        btnRestartNow.TabIndex = 2;
        btnRestartNow.Text = "Reiniciar agora";
        btnRestartNow.UseVisualStyleBackColor = true;
        btnRestartNow.Click += btnRestartNow_Click;
        // 
        // btnKillExcel
        // 
        btnKillExcel.AutoSize = true;
        btnKillExcel.Location = new Point(351, 3);
        btnKillExcel.Name = "btnKillExcel";
        btnKillExcel.Size = new Size(110, 35);
        btnKillExcel.TabIndex = 3;
        btnKillExcel.Text = "Matar Excel";
        btnKillExcel.UseVisualStyleBackColor = true;
        btnKillExcel.Click += btnKillExcel_Click;
        // 
        // btnOpenLogFolder
        // 
        btnOpenLogFolder.AutoSize = true;
        btnOpenLogFolder.Location = new Point(467, 3);
        btnOpenLogFolder.Name = "btnOpenLogFolder";
        btnOpenLogFolder.Size = new Size(150, 35);
        btnOpenLogFolder.TabIndex = 4;
        btnOpenLogFolder.Text = "Abrir pasta do log";
        btnOpenLogFolder.UseVisualStyleBackColor = true;
        btnOpenLogFolder.Click += btnOpenLogFolder_Click;
        // 
        // lblStatusTitle
        // 
        lblStatusTitle.AutoSize = true;
        lblStatusTitle.Dock = DockStyle.Fill;
        lblStatusTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblStatusTitle.Location = new Point(3, 0);
        lblStatusTitle.Name = "lblStatusTitle";
        lblStatusTitle.Size = new Size(154, 30);
        lblStatusTitle.TabIndex = 0;
        lblStatusTitle.Text = "Status";
        lblStatusTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblStatusValue
        // 
        lblStatusValue.AutoEllipsis = true;
        lblStatusValue.AutoSize = true;
        lblStatusValue.Dock = DockStyle.Fill;
        lblStatusValue.Location = new Point(163, 0);
        lblStatusValue.Name = "lblStatusValue";
        lblStatusValue.Size = new Size(780, 30);
        lblStatusValue.TabIndex = 1;
        lblStatusValue.Text = "-";
        lblStatusValue.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblLastExecutionTitle
        // 
        lblLastExecutionTitle.AutoSize = true;
        lblLastExecutionTitle.Dock = DockStyle.Fill;
        lblLastExecutionTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblLastExecutionTitle.Location = new Point(3, 30);
        lblLastExecutionTitle.Name = "lblLastExecutionTitle";
        lblLastExecutionTitle.Size = new Size(154, 30);
        lblLastExecutionTitle.TabIndex = 2;
        lblLastExecutionTitle.Text = "Última execução";
        lblLastExecutionTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblLastExecutionValue
        // 
        lblLastExecutionValue.AutoEllipsis = true;
        lblLastExecutionValue.AutoSize = true;
        lblLastExecutionValue.Dock = DockStyle.Fill;
        lblLastExecutionValue.Location = new Point(163, 30);
        lblLastExecutionValue.Name = "lblLastExecutionValue";
        lblLastExecutionValue.Size = new Size(780, 30);
        lblLastExecutionValue.TabIndex = 3;
        lblLastExecutionValue.Text = "-";
        lblLastExecutionValue.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblHtmlLastUpdateTitle
        // 
        lblHtmlLastUpdateTitle.AutoSize = true;
        lblHtmlLastUpdateTitle.Dock = DockStyle.Fill;
        lblHtmlLastUpdateTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblHtmlLastUpdateTitle.Location = new Point(3, 60);
        lblHtmlLastUpdateTitle.Name = "lblHtmlLastUpdateTitle";
        lblHtmlLastUpdateTitle.Size = new Size(154, 30);
        lblHtmlLastUpdateTitle.TabIndex = 4;
        lblHtmlLastUpdateTitle.Text = "Última atualização HTML";
        lblHtmlLastUpdateTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblHtmlLastUpdateValue
        // 
        lblHtmlLastUpdateValue.AutoEllipsis = true;
        lblHtmlLastUpdateValue.AutoSize = true;
        lblHtmlLastUpdateValue.Dock = DockStyle.Fill;
        lblHtmlLastUpdateValue.Location = new Point(163, 60);
        lblHtmlLastUpdateValue.Name = "lblHtmlLastUpdateValue";
        lblHtmlLastUpdateValue.Size = new Size(780, 30);
        lblHtmlLastUpdateValue.TabIndex = 5;
        lblHtmlLastUpdateValue.Text = "-";
        lblHtmlLastUpdateValue.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblHtmlPathTitle
        // 
        lblHtmlPathTitle.AutoSize = true;
        lblHtmlPathTitle.Dock = DockStyle.Fill;
        lblHtmlPathTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblHtmlPathTitle.Location = new Point(3, 90);
        lblHtmlPathTitle.Name = "lblHtmlPathTitle";
        lblHtmlPathTitle.Size = new Size(154, 30);
        lblHtmlPathTitle.TabIndex = 6;
        lblHtmlPathTitle.Text = "HTML monitorado";
        lblHtmlPathTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblHtmlPathValue
        // 
        lblHtmlPathValue.AutoEllipsis = true;
        lblHtmlPathValue.AutoSize = true;
        lblHtmlPathValue.Dock = DockStyle.Fill;
        lblHtmlPathValue.Location = new Point(163, 90);
        lblHtmlPathValue.Name = "lblHtmlPathValue";
        lblHtmlPathValue.Size = new Size(780, 30);
        lblHtmlPathValue.TabIndex = 7;
        lblHtmlPathValue.Text = "-";
        lblHtmlPathValue.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblNextActionTitle
        // 
        lblNextActionTitle.AutoSize = true;
        lblNextActionTitle.Dock = DockStyle.Fill;
        lblNextActionTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblNextActionTitle.Location = new Point(3, 120);
        lblNextActionTitle.Name = "lblNextActionTitle";
        lblNextActionTitle.Size = new Size(154, 30);
        lblNextActionTitle.TabIndex = 8;
        lblNextActionTitle.Text = "Próxima ação";
        lblNextActionTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblNextActionValue
        // 
        lblNextActionValue.AutoEllipsis = true;
        lblNextActionValue.AutoSize = true;
        lblNextActionValue.Dock = DockStyle.Fill;
        lblNextActionValue.Location = new Point(163, 120);
        lblNextActionValue.Name = "lblNextActionValue";
        lblNextActionValue.Size = new Size(780, 30);
        lblNextActionValue.TabIndex = 9;
        lblNextActionValue.Text = "-";
        lblNextActionValue.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblCurrentMessageTitle
        // 
        lblCurrentMessageTitle.AutoSize = true;
        lblCurrentMessageTitle.Dock = DockStyle.Fill;
        lblCurrentMessageTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblCurrentMessageTitle.Location = new Point(3, 150);
        lblCurrentMessageTitle.Name = "lblCurrentMessageTitle";
        lblCurrentMessageTitle.Size = new Size(154, 70);
        lblCurrentMessageTitle.TabIndex = 10;
        lblCurrentMessageTitle.Text = "Mensagem atual";
        lblCurrentMessageTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // txtCurrentMessage
        // 
        txtCurrentMessage.Dock = DockStyle.Fill;
        txtCurrentMessage.Location = new Point(163, 153);
        txtCurrentMessage.Multiline = true;
        txtCurrentMessage.Name = "txtCurrentMessage";
        txtCurrentMessage.ReadOnly = true;
        txtCurrentMessage.ScrollBars = ScrollBars.Vertical;
        txtCurrentMessage.Size = new Size(780, 64);
        txtCurrentMessage.TabIndex = 11;
        // 
        // txtLogs
        // 
        txtLogs.Dock = DockStyle.Fill;
        txtLogs.Font = new Font("Consolas", 9F);
        txtLogs.Location = new Point(15, 355);
        txtLogs.Multiline = true;
        txtLogs.Name = "txtLogs";
        txtLogs.ReadOnly = true;
        txtLogs.ScrollBars = ScrollBars.Vertical;
        txtLogs.Size = new Size(946, 271);
        txtLogs.TabIndex = 3;
        // 
        // lblLogsTitle
        // 
        lblLogsTitle.AutoSize = true;
        lblLogsTitle.Dock = DockStyle.Fill;
        lblLogsTitle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        lblLogsTitle.Location = new Point(15, 327);
        lblLogsTitle.Name = "lblLogsTitle";
        lblLogsTitle.Size = new Size(946, 28);
        lblLogsTitle.TabIndex = 2;
        lblLogsTitle.Text = "Mensagens recentes de log";
        lblLogsTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // tableMain
        // 
        tableMain.ColumnCount = 1;
        tableMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableMain.Controls.Add(panelButtons, 0, 0);
        tableMain.Controls.Add(tableStatus, 0, 1);
        tableMain.Controls.Add(lblLogsTitle, 0, 2);
        tableMain.Controls.Add(txtLogs, 0, 3);
        tableMain.Dock = DockStyle.Fill;
        tableMain.Location = new Point(0, 0);
        tableMain.Name = "tableMain";
        tableMain.Padding = new Padding(12);
        tableMain.RowCount = 4;
        tableMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 48F));
        tableMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 264F));
        tableMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
        tableMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        tableMain.Size = new Size(976, 641);
        tableMain.TabIndex = 0;
        // 
        // panelButtons
        // 
        panelButtons.Controls.Add(btnStart);
        panelButtons.Controls.Add(btnStop);
        panelButtons.Controls.Add(btnRestartNow);
        panelButtons.Controls.Add(btnKillExcel);
        panelButtons.Controls.Add(btnOpenLogFolder);
        panelButtons.Dock = DockStyle.Fill;
        panelButtons.Location = new Point(15, 15);
        panelButtons.Name = "panelButtons";
        panelButtons.Size = new Size(946, 42);
        panelButtons.TabIndex = 0;
        // 
        // tableStatus
        // 
        tableStatus.ColumnCount = 2;
        tableStatus.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
        tableStatus.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableStatus.Controls.Add(lblStatusTitle, 0, 0);
        tableStatus.Controls.Add(lblStatusValue, 1, 0);
        tableStatus.Controls.Add(lblLastExecutionTitle, 0, 1);
        tableStatus.Controls.Add(lblLastExecutionValue, 1, 1);
        tableStatus.Controls.Add(lblHtmlLastUpdateTitle, 0, 2);
        tableStatus.Controls.Add(lblHtmlLastUpdateValue, 1, 2);
        tableStatus.Controls.Add(lblHtmlPathTitle, 0, 3);
        tableStatus.Controls.Add(lblHtmlPathValue, 1, 3);
        tableStatus.Controls.Add(lblNextActionTitle, 0, 4);
        tableStatus.Controls.Add(lblNextActionValue, 1, 4);
        tableStatus.Controls.Add(lblCurrentMessageTitle, 0, 5);
        tableStatus.Controls.Add(txtCurrentMessage, 1, 5);
        tableStatus.Controls.Add(lblConfigPathTitle, 0, 6);
        tableStatus.Controls.Add(lblConfigPathValue, 1, 6);
        tableStatus.Controls.Add(lblLogFolderTitle, 0, 7);
        tableStatus.Controls.Add(lblLogFolderValue, 1, 7);
        tableStatus.Dock = DockStyle.Fill;
        tableStatus.Location = new Point(15, 63);
        tableStatus.Name = "tableStatus";
        tableStatus.RowCount = 8;
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableStatus.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        tableStatus.Size = new Size(946, 258);
        tableStatus.TabIndex = 1;
        // 
        // lblConfigPathTitle
        // 
        lblConfigPathTitle.AutoSize = true;
        lblConfigPathTitle.Dock = DockStyle.Fill;
        lblConfigPathTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblConfigPathTitle.Location = new Point(3, 220);
        lblConfigPathTitle.Name = "lblConfigPathTitle";
        lblConfigPathTitle.Size = new Size(154, 30);
        lblConfigPathTitle.TabIndex = 12;
        lblConfigPathTitle.Text = "appsettings.json";
        lblConfigPathTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblConfigPathValue
        // 
        lblConfigPathValue.AutoEllipsis = true;
        lblConfigPathValue.AutoSize = true;
        lblConfigPathValue.Dock = DockStyle.Fill;
        lblConfigPathValue.Location = new Point(163, 220);
        lblConfigPathValue.Name = "lblConfigPathValue";
        lblConfigPathValue.Size = new Size(780, 30);
        lblConfigPathValue.TabIndex = 13;
        lblConfigPathValue.Text = "-";
        lblConfigPathValue.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblLogFolderTitle
        // 
        lblLogFolderTitle.AutoSize = true;
        lblLogFolderTitle.Dock = DockStyle.Fill;
        lblLogFolderTitle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblLogFolderTitle.Location = new Point(3, 250);
        lblLogFolderTitle.Name = "lblLogFolderTitle";
        lblLogFolderTitle.Size = new Size(154, 30);
        lblLogFolderTitle.TabIndex = 14;
        lblLogFolderTitle.Text = "Pasta de log";
        lblLogFolderTitle.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // lblLogFolderValue
        // 
        lblLogFolderValue.AutoEllipsis = true;
        lblLogFolderValue.AutoSize = true;
        lblLogFolderValue.Dock = DockStyle.Fill;
        lblLogFolderValue.Location = new Point(163, 250);
        lblLogFolderValue.Name = "lblLogFolderValue";
        lblLogFolderValue.Size = new Size(780, 30);
        lblLogFolderValue.TabIndex = 15;
        lblLogFolderValue.Text = "-";
        lblLogFolderValue.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // MainForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(976, 641);
        Controls.Add(tableMain);
        MinimumSize = new Size(900, 600);
        Name = "MainForm";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Excel Macro Runner";
        tableMain.ResumeLayout(false);
        tableMain.PerformLayout();
        panelButtons.ResumeLayout(false);
        panelButtons.PerformLayout();
        tableStatus.ResumeLayout(false);
        tableStatus.PerformLayout();
        ResumeLayout(false);
    }
}
