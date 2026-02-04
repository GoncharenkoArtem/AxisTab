using AxisTAb;
using System;
using System.Collections.Generic;

namespace AxisTAb
{
    public static class DrawingHost
    {
        public static readonly Dictionary<IntPtr, DrawingAxis> drawingDictionary = new Dictionary<IntPtr, DrawingAxis>();

        public static DrawingAxis Current
        {
            get
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc is null) throw new InvalidOperationException("Нет активного документа.");
                var key = doc.UnmanagedObject; // стабильный ключ для документа

                if (!drawingDictionary.TryGetValue(key, out var s))
                {
                    s = new DrawingAxis();
                    s.doc = doc;
                    s.axis = RibbonInitializer.GetListOfAxis(); // заполняем List<Axis> axis 
                    if (s.axis.Count > 0) s.currentAxis = s.axis[0];    // текущая ось
                    drawingDictionary[key] = s;
                }
                return s;

            }
        }
    }
}
