using System.Windows.Forms;

namespace CADVISER_Transcript_Customers
{
    partial class Default
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Try_CADVISER_DB = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Test_Data = new System.Windows.Forms.Button();
            this.Try_Temp_DB = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.Log = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.Get_Data = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // Try_CADVISER_DB
            // 
            this.Try_CADVISER_DB.Location = new System.Drawing.Point(10, 21);
            this.Try_CADVISER_DB.Name = "Try_CADVISER_DB";
            this.Try_CADVISER_DB.Size = new System.Drawing.Size(106, 23);
            this.Try_CADVISER_DB.TabIndex = 0;
            this.Try_CADVISER_DB.Text = "鑫全城資料庫";
            this.Try_CADVISER_DB.UseVisualStyleBackColor = true;
            this.Try_CADVISER_DB.Click += new System.EventHandler(this.Try_CADVISER_DB_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Test_Data);
            this.groupBox1.Controls.Add(this.Try_Temp_DB);
            this.groupBox1.Controls.Add(this.Try_CADVISER_DB);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(122, 107);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "測試連接";
            // 
            // Test_Data
            // 
            this.Test_Data.Location = new System.Drawing.Point(10, 79);
            this.Test_Data.Name = "Test_Data";
            this.Test_Data.Size = new System.Drawing.Size(106, 23);
            this.Test_Data.TabIndex = 2;
            this.Test_Data.Text = "測試";
            this.Test_Data.UseVisualStyleBackColor = true;
            this.Test_Data.Click += new System.EventHandler(this.Test_Data_Click);
            // 
            // Try_Temp_DB
            // 
            this.Try_Temp_DB.Location = new System.Drawing.Point(10, 50);
            this.Try_Temp_DB.Name = "Try_Temp_DB";
            this.Try_Temp_DB.Size = new System.Drawing.Size(106, 23);
            this.Try_Temp_DB.TabIndex = 1;
            this.Try_Temp_DB.Text = "Temp資料庫";
            this.Try_Temp_DB.UseVisualStyleBackColor = true;
            this.Try_Temp_DB.Click += new System.EventHandler(this.Try_Temp_DB_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.Log);
            this.groupBox2.Location = new System.Drawing.Point(12, 234);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(662, 123);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "運算結果";
            // 
            // Log
            // 
            this.Log.Location = new System.Drawing.Point(6, 21);
            this.Log.Multiline = true;
            this.Log.Name = "Log";
            this.Log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.Log.Size = new System.Drawing.Size(650, 96);
            this.Log.TabIndex = 3;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.Get_Data);
            this.groupBox3.Location = new System.Drawing.Point(165, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(122, 82);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "資料匯出_功能調適";
            // 
            // Get_Data
            // 
            this.Get_Data.Location = new System.Drawing.Point(10, 21);
            this.Get_Data.Name = "Get_Data";
            this.Get_Data.Size = new System.Drawing.Size(106, 23);
            this.Get_Data.TabIndex = 0;
            this.Get_Data.Text = "取得資料";
            this.Get_Data.UseVisualStyleBackColor = true;
            this.Get_Data.Click += new System.EventHandler(this.Get_Data_Click);
            // 
            // Default
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(686, 369);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Default";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Try_CADVISER_DB;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button Try_Temp_DB;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox Log;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button Get_Data;
        private System.Windows.Forms.Button Test_Data;
    }
}

