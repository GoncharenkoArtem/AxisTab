using ACAD_test;
using System;
using System.Collections.Generic;

namespace TSODD
{
    public static class TsoddHost
    {
        public static readonly Dictionary<IntPtr, DrawingTSODD> tsoddDictionary = new Dictionary<IntPtr, DrawingTSODD>();

        public static DrawingTSODD Current
        {
            get
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc is null) throw new InvalidOperationException("Нет активного документа.");
                var key = doc.UnmanagedObject; // стабильный ключ для документа

                if (!tsoddDictionary.TryGetValue(key, out var s))
                {
                    s = new DrawingTSODD();
                    s.axis = RibbonInitializer.GetListOfAxis(); // заполняем List<Axis> axis 
                    if (s.axis.Count > 0) s.currentAxis = s.axis[0];    // текущая ось
                    tsoddDictionary[key] = s;
                }
                return s;

            }
        }
    }
}
