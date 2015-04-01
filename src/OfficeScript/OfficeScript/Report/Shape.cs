﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.Report
{
    class Shape
    {
        private PowerPoint.Shape shape;
        private bool disposed;

        public Shape(PowerPoint.Shape shape)
        {
            this.shape = shape;
        }

        public object Invoke()
        {
            return new
            {
                attr = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return Util.Attr(this, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value), Invoke);
                    }),
                tags = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return new Tags(this.shape).Invoke();
                    }),
                remove = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Remove();
                        return null;
                    })
            };
        }


        /// <summary>
        /// Deletes the Shape
        /// </summary>
        private void Remove()
        {
            this.shape.Delete();
            this.shape.Dispose();
        }

        #region Properties

        public string Name
        {
            get
            {
                return this.shape.Name;
            }
            set
            {
                this.shape.Name = value;
            }
        }
        public string Text
        {
            get
            {
                
                return this.shape.TextFrame.TextRange.Text;
                
            }
            set
            {
                this.shape.TextFrame.TextRange.Text = value;
               
            }
        }

        /// <summary>
        /// Get or Set the Top-Property for this element.
        /// </summary>
        public float Top
        {
            get
            {
                return this.shape.Top;
            }
            set
            {
                this.shape.Top = value;
            }
        }
        /// <summary>
        /// Get or Set the Left-Property for this element.
        /// </summary>
        public float Left
        {
            get
            {
                return this.shape.Left;
            }
            set
            {
                this.shape.Left = value;
            }
        }
        /// <summary>
        /// Get or Set the Height-Property for this element.
        /// </summary>
        public float Height
        {
            get
            {
                return this.shape.Height;
            }
            set
            {
                this.shape.Height = value;
            }
        }
        /// <summary>
        /// Get or Set the Width-Property for this element.
        /// </summary>
        public float Width
        {
            get
            {
                return this.shape.Width;
            }
            set
            {
                this.shape.Width = value;
            }
        }
        /// <summary>
        /// Get or Set the Rotation-Property for this element.
        /// </summary>
        public float Rotation
        {
            get
            {
                return this.shape.Rotation;
            }
            set
            {
                this.shape.Rotation = value;
            }
        }

        #endregion Properties
    }
}