﻿/*
    The Insulator creates a proper, dynamic Insulation in Autodesk's (R) Revit (R)
    Copyright (C) 2014  Maximilian Thumfart

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Insulator
{
    public partial class InsulatorMaterial : Form
    {
        public InsulatorMaterial()
        {
            InitializeComponent();
        }

        public bool zigzag = false;

        private void button1_Click(object sender, EventArgs e)
        {
            this.zigzag = false;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.zigzag = true;
            this.Close();
        }
    }
}
