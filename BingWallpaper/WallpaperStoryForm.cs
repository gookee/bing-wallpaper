﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BingWallpaper
{
    public partial class WallpaperStoryForm : Form
    {
        HistoryImage image;

        public WallpaperStoryForm(HistoryImage image)
        {
            InitializeComponent();
            this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            this.panel1.BackColor = Color.FromArgb(100, 100, 100, 100);
            this.panel1.Parent = this.pictureBox1;
            this.btnNext.Parent = this.pictureBox1;
            this.btnPre.Parent = this.pictureBox1;
            this.btnPre.Top = this.btnNext.Top = this.Height / 2 - 20;
            this.image = image;
        }

        private void WallpaperStoryForm_LoadAsync(object sender, EventArgs e)
        {
            LoadImageAsync();
        }

        async void LoadImageAsync()
        {
            if (image != null)
            {
                this.lblDate.Text = image.Date;
                this.lblDesc.Text = image.Description;
                this.lblLocation.Text = image.Locate;
                this.Text = image.Title;
                this.lblTitle.Text = image.Title;

                try
                {
                    var img = await image.getImage();
                    pictureBox1.Image = img;
                }
                catch
                {
                    pictureBox1.Image = Properties.Resources.error;
                }

            }
        }

        private void btnPre_Click(object sender, EventArgs e)
        {
            JumpPre();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            JumpNext();
        }

        private void WallpaperStoryForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Left)
            {
                JumpPre();
            }
            else if (e.KeyCode == Keys.Right)
            {
                JumpNext();
            }
        }

        private void JumpPre()
        {
            var pre = HistoryImageProvider.Previous(image.Date);
            if (pre != null)
            {
                image = pre;
            }
            else
            {
                image = HistoryImageProvider.Last();
            }
            LoadImageAsync();
        }

        private void JumpNext()
        {
            var next = HistoryImageProvider.Next(image.Date);
            if (next != null)
            {
                image = next;
            }
            else
            {
                image = HistoryImageProvider.First();
            }
            LoadImageAsync();
        }
    }
}
