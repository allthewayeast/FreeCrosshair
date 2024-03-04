using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace FreeCrosshair
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                using (var app = new ApplicationBase())
                {
                    app.Run(args);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Extensions
        public static string ToHexString(this Color source)
        {
            return $"{source.A:X2}{source.R:X2}{source.G:X2}{source.B:X2}";
        }

        public static T NextOrDefault<T>(this IEnumerable<T> source, T current)
        {
            using (var enumerator = source.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current.Equals(current))
                    {
                        if (enumerator.MoveNext())
                        {
                            return enumerator.Current;
                        }
                    }
                }
            }

            return default(T);
        }
        #endregion
    }
}
