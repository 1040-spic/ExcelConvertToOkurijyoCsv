using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelConvertToOkumarukunnCsv.Forms
{
	partial class OutputOkumarukunnCsvFrm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        /// 
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.JyutyuDropArea = new System.Windows.Forms.Label();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.OkurijyoDropArea = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Outset;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.JyutyuDropArea, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(62, 380);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(10);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 200F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(622, 200);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // JyutyuDropArea
            // 
            this.JyutyuDropArea.AllowDrop = true;
            this.JyutyuDropArea.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(251)))), ((int)(((byte)(254)))));
            this.JyutyuDropArea.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.JyutyuDropArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.JyutyuDropArea.Font = new System.Drawing.Font("メイリオ", 12F);
            this.JyutyuDropArea.Location = new System.Drawing.Point(5, -2);
            this.JyutyuDropArea.Name = "JyutyuDropArea";
            this.JyutyuDropArea.Padding = new System.Windows.Forms.Padding(10);
            this.JyutyuDropArea.Size = new System.Drawing.Size(612, 200);
            this.JyutyuDropArea.TabIndex = 1;
            this.JyutyuDropArea.Text = "こちらに受注一覧ファイルをドロップ";
            this.JyutyuDropArea.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.JyutyuDropArea.DragDrop += new System.Windows.Forms.DragEventHandler(this.JyutyuDropArea_DragDrop);
            this.JyutyuDropArea.DragEnter += new System.Windows.Forms.DragEventHandler(this.JyutyuDropArea_DragEnter);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Outset;
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Controls.Add(this.OkurijyoDropArea, 0, 1);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(67, 100);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(10);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 200F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(612, 200);
            this.tableLayoutPanel2.TabIndex = 1;
            // 
            // OkurijyoDropArea
            // 
            this.OkurijyoDropArea.AllowDrop = true;
            this.OkurijyoDropArea.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(181)))), ((int)(((byte)(232)))), ((int)(((byte)(221)))));
            this.OkurijyoDropArea.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.OkurijyoDropArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.OkurijyoDropArea.Font = new System.Drawing.Font("メイリオ", 12F);
            this.OkurijyoDropArea.Location = new System.Drawing.Point(5, -2);
            this.OkurijyoDropArea.Name = "OkurijyoDropArea";
            this.OkurijyoDropArea.Padding = new System.Windows.Forms.Padding(10);
            this.OkurijyoDropArea.Size = new System.Drawing.Size(602, 200);
            this.OkurijyoDropArea.TabIndex = 1;
            this.OkurijyoDropArea.Text = "こちらに送り状ファイルをドロップ";
            this.OkurijyoDropArea.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.OkurijyoDropArea.DragDrop += new System.Windows.Forms.DragEventHandler(this.OkurijyoDropArea_DragDrop);
            this.OkurijyoDropArea.DragEnter += new System.Windows.Forms.DragEventHandler(this.OkurijyoDropArea_DragEnter);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("メイリオ", 12F);
            this.label1.Location = new System.Drawing.Point(56, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(423, 36);
            this.label1.TabIndex = 5;
            this.label1.Text = "【送り状出荷データファイルの変換】";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("メイリオ", 12F);
            this.label2.Location = new System.Drawing.Point(56, 319);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(327, 36);
            this.label2.TabIndex = 6;
            this.label2.Text = "【受注一覧ファイルの変換】";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("メイリオ", 8F);
            this.label3.Location = new System.Drawing.Point(68, 74);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(602, 24);
            this.label3.TabIndex = 7;
            this.label3.Text = "ファイル名に\'出荷\'が含まれるExcelファイルを以下エリアにドロップしてください";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("メイリオ", 8F);
            this.label4.Location = new System.Drawing.Point(63, 354);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(602, 24);
            this.label4.TabIndex = 8;
            this.label4.Text = "ファイル名に\'受注\'が含まれるExcelファイルを以下エリアにドロップしてください";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // OutputOkumarukunnCsvFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(745, 628);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Margin = new System.Windows.Forms.Padding(10);
            this.Name = "OutputOkumarukunnCsvFrm";
            this.Text = "『おくまるくん』用データ変換CSV出力";
            this.Resize += new System.EventHandler(this.OutputOkumarukunnCsvFrm_Resize);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void OutputOkumarukunnCsvFrm_Resize(object sender, EventArgs e)
        {
            // フォームの幅に応じて文字サイズを調整
            float widthFactor = (float)this.ClientSize.Width / 745f; // 基準となるフォーム幅を使用
            float newFontSizeLabel1and2 = 16f * widthFactor; // label1 と label2 の文字サイズ調整
            float newFontSizeLabel3and4 = 12f * widthFactor; // label3 と label4 の文字サイズ調整

            // label1, label2, label3, label4のフォントサイズを変更
            this.label1.Font = new Font(this.label1.Font.FontFamily, newFontSizeLabel1and2);
            this.label2.Font = new Font(this.label2.Font.FontFamily, newFontSizeLabel1and2);
            this.label3.Font = new Font(this.label3.Font.FontFamily, newFontSizeLabel3and4);
            this.label4.Font = new Font(this.label4.Font.FontFamily, newFontSizeLabel3and4);
        }

        private void lblDropArea_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        #endregion
        
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label OkurijyoDropArea;
        private System.Windows.Forms.Label JyutyuDropArea;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
    }
}