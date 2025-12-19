using System.Reflection;

namespace Escala
{
    /// <summary>
    /// Métodos de extensão para controles Windows Forms
    /// </summary>
    public static class ExtensionMethods
    {
        /// <summary>
        /// Habilita ou desabilita double buffering em um DataGridView para melhorar performance
        /// </summary>
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null)
                pi.SetValue(dgv, setting, null);
        }
    }
}
