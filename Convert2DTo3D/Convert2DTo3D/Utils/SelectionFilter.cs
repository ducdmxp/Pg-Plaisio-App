using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D.Utils
{
    public class SelectionFilter
    {
    }

    public class TypeSelectionFilter : ISelectionFilter
    {
        private List<Category> _categories;

        private RevitLinkInstance m_currentInstance = null;

        public TypeSelectionFilter(List<Category> categories = null)
        {
            _categories = categories;
        }

        public bool AllowElement(Element elem)
        {
            if (elem == null || elem.Category == null)
                return false;

            if (elem is RevitLinkInstance link)
            {
                m_currentInstance = link;
            }
            else
            {
                if (_categories != null && !_categories.Select(x => x.Id).Contains(elem.Category.Id))
                    return false;
            }

            return true;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            if (m_currentInstance == null)
                return false;

            Document linkedDoc = m_currentInstance.GetLinkDocument();
            Element elem = linkedDoc.GetElement(reference.LinkedElementId);

            if (elem != null && elem.Category != null)
            {
                if (_categories != null && !_categories.Select(x => x.Id).Contains(elem.Category.Id))
                    return false;

                return true;
            }

            return false;
        }
    }

    public class PipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Pipe;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}