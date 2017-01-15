namespace Importer
{
    partial class MainForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.відкритиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.вихідToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.налаштуванняToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.прайсДляСайтаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.кросиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.конвертаціяToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.замінаБрендівToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.налаштуванняToolStripMenuItem,
            this.прайсДляСайтаToolStripMenuItem,
            this.кросиToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1107, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.відкритиToolStripMenuItem,
            this.вихідToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // відкритиToolStripMenuItem
            // 
            this.відкритиToolStripMenuItem.Name = "відкритиToolStripMenuItem";
            this.відкритиToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.відкритиToolStripMenuItem.Text = "Конвертувати один файл";
            this.відкритиToolStripMenuItem.Click += new System.EventHandler(this.відкритиToolStripMenuItem_Click);
            // 
            // вихідToolStripMenuItem
            // 
            this.вихідToolStripMenuItem.Name = "вихідToolStripMenuItem";
            this.вихідToolStripMenuItem.Size = new System.Drawing.Size(119, 22);
            this.вихідToolStripMenuItem.Text = "Вихід";
            this.вихідToolStripMenuItem.Click += new System.EventHandler(this.вихідToolStripMenuItem_Click);
            // 
            // налаштуванняToolStripMenuItem
            // 
            this.налаштуванняToolStripMenuItem.Name = "налаштуванняToolStripMenuItem";
            this.налаштуванняToolStripMenuItem.Size = new System.Drawing.Size(94, 20);
            this.налаштуванняToolStripMenuItem.Text = "Налаштування";
            this.налаштуванняToolStripMenuItem.Click += new System.EventHandler(this.налаштуванняToolStripMenuItem_Click);
            // 
            // прайсДляСайтаToolStripMenuItem
            // 
            this.прайсДляСайтаToolStripMenuItem.Name = "прайсДляСайтаToolStripMenuItem";
            this.прайсДляСайтаToolStripMenuItem.Size = new System.Drawing.Size(122, 20);
            this.прайсДляСайтаToolStripMenuItem.Text = "Конвертація прайсів";
            this.прайсДляСайтаToolStripMenuItem.Click += new System.EventHandler(this.прайсДляСайтаToolStripMenuItem_Click);
            // 
            // кросиToolStripMenuItem
            // 
            this.кросиToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.конвертаціяToolStripMenuItem,
            this.замінаБрендівToolStripMenuItem1});
            this.кросиToolStripMenuItem.Name = "кросиToolStripMenuItem";
            this.кросиToolStripMenuItem.Size = new System.Drawing.Size(49, 20);
            this.кросиToolStripMenuItem.Text = "Кроси";
            this.кросиToolStripMenuItem.Click += new System.EventHandler(this.кросиToolStripMenuItem_Click);
            // 
            // конвертаціяToolStripMenuItem
            // 
            this.конвертаціяToolStripMenuItem.Name = "конвертаціяToolStripMenuItem";
            this.конвертаціяToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.конвертаціяToolStripMenuItem.Text = "Конвертація";
            this.конвертаціяToolStripMenuItem.Click += new System.EventHandler(this.конвертаціяToolStripMenuItem_Click);
            // 
            // замінаБрендівToolStripMenuItem1
            // 
            this.замінаБрендівToolStripMenuItem1.Name = "замінаБрендівToolStripMenuItem1";
            this.замінаБрендівToolStripMenuItem1.Size = new System.Drawing.Size(148, 22);
            this.замінаБрендівToolStripMenuItem1.Text = "Заміна брендів";
            this.замінаБрендівToolStripMenuItem1.Click += new System.EventHandler(this.замінаБрендівToolStripMenuItem1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1107, 425);
            this.Controls.Add(this.menuStrip1);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "Конвертер прайсів";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem відкритиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem вихідToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem налаштуванняToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem прайсДляСайтаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem кросиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem конвертаціяToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem замінаБрендівToolStripMenuItem1;
    }
}

