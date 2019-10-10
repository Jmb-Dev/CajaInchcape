using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace CajaIndigo
{
  public class PDFPageNumber
    {
      private PdfTemplate totalPages;

      //public void onOpenDocument(PdfWriter writer, Document document)
      //{
      //    totalPages = writer.DirectContent.CreateTemplate(100, 100);
      //    totalPages.BoundingBox =  new Rectangle(-20, -20, 100, 100);
      //}

      protected PdfTemplate total;
      protected BaseFont helv;
      private bool settingFont = false;

      public void OnOpenDocument(PdfWriter writer, Document document)
      {
          total = writer.DirectContent.CreateTemplate(100, 100);
          total.BoundingBox = new Rectangle(-20, -20, 100, 100);

          helv = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
      }

      public void OnEndPage(PdfWriter writer, Document document)
      {
          helv = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
          total = writer.DirectContent.CreateTemplate(100, 100);
          total.BoundingBox = new Rectangle(-20, -20, 100, 100);

          PdfContentByte cb = writer.DirectContent;
          cb.SaveState();
          string text = "Page " + writer.PageNumber + " of ";
          float textBase = document.Bottom - 20;
          float textSize = 12; //helv.GetWidthPoint(text, 12);
          cb.BeginText();
          cb.SetFontAndSize(helv, 12);
          if ((writer.PageNumber % 2) == 1)
          {
              cb.SetTextMatrix(document.Left, textBase);
              cb.ShowText(text);
              cb.EndText();
              cb.AddTemplate(total, document.Left + textSize, textBase);
          }
          else
          {
              float adjust = helv.GetWidthPoint("0", 12);
              cb.SetTextMatrix(document.Right - textSize - adjust, textBase);
              cb.ShowText(text);
              cb.EndText();
              cb.AddTemplate(total, document.Right - adjust, textBase);
          }
          cb.RestoreState();
      }

      public void OnCloseDocument(PdfWriter writer, Document document)
      {
          total.BeginText();
          total.SetFontAndSize(helv, 12);
          total.SetTextMatrix(0, 0);
          int pageNumber = writer.PageNumber - 1;
          total.ShowText(Convert.ToString(pageNumber));
          total.EndText();
      }

    }
}
