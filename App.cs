// ============================================================
//  ПОМОЩНИК ПО ЗАЯВЛЕНИЯМ
//  1. Исправление документов (реквизиты, КПП, кавычки, копейки)
//  2. Прописание сумм: цифры из Word -> пропись в скобках
// ============================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ZayavleniyaApp
{
    // ---------- параметры генератора ----------
    public class GenParams
    {
        public string NumPrefix = "ВС№";
        public string Num;
        public string Date;
        public string Addr;
        public string Rep = "";
        public bool Solidary;
        public List<string> Debtors = new List<string>();
        public string Osn, Gos, Sud;
        public bool UseOsn = true, UseGos = true, UseSud = true;

        public string FullNum
        {
            get
            {
                string n = (Num ?? "").Trim();
                string p = (NumPrefix ?? "").Trim();
                if (string.IsNullOrEmpty(p)) return n;
                if (n.StartsWith(p, StringComparison.OrdinalIgnoreCase)) return n;
                return p + " " + n;
            }
        }

        public static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Заявление";
            char[] invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (char c in name.Trim())
            {
                if (c == '/' || c == '\\') sb.Append('_');
                else if (invalid.Contains(c)) sb.Append('_');
                else sb.Append(c);
            }
            string res = sb.ToString().Trim(' ', '.');
            return string.IsNullOrEmpty(res) ? "Заявление" : res;
        }
    }

    public partial class Engine
    {
        #region Создание стандартного шаблона (Встроенный эталон 094729862.docx)
        // Встроенный эталонный образец заявления (094729862.docx)
        const string EmbeddedMasterTemplateB64 =
            "UEsDBBQABgAIAAAAIQA+UkjocQEAAKQFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0lMtqwzAQRfeF/oPRtsRKuiilxMmij2UbaArdKvI4EdULafL6+47txJTixKVJNgZ55p57NYgZjjdGJysI" +
            "UTmbsUHaZwlY6XJl5xn7mL707lkSUdhcaGchY1uIbDy6vhpOtx5iQmobM7ZA9A+cR7kAI2LqPFiqFC4YgXQMc+6F/BJz4Lf9/h2XziJY7GHJYKPhExRiqTF5" +
            "3tDvOkkAHVnyWDeWXhkT3mslBVKdr2z+y6W3c0hJWfXEhfLxhhoYb3UoK4cNdro3Gk1QOSQTEfBVGOriaxdynju5NKRMj2NacrqiUBIafUnzwUmIkWZudNpU" +
            "jFB2n78th1xGdObTaK4QzCQ4Hwcnx2mgJQ8CKmhmeHAWEbca4vknUXO77QGRBJcIsCN3RljD7P1iKX7AO4MU5DsVMw3nj9GgO0MgbQGov6c/yApzzJI6q7dP" +
            "WyX849r7tVGqe/5Pj75xJPTJ94NyI+WQt3jzaseOvgEAAP//AwBQSwMEFAAGAAgAAAAhAB6RGrfvAAAATgIAAAsACAJfcmVscy8ucmVscyCiBAIooAACAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsksFqwzAMQO+D/YPRvVHawRijTi9j0NsY2QcIW0lM" +
            "E9vYatf+/TzY2AJd6WFHy9LTk9B6c5xGdeCUXfAallUNir0J1vlew1v7vHgAlYW8pTF41nDiDJvm9mb9yiNJKcqDi1kVis8aBpH4iJjNwBPlKkT25acLaSIp" +
            "z9RjJLOjnnFV1/eYfjOgmTHV1mpIW3sHqj1FvoYdus4ZfgpmP7GXMy2Qj8Lesl3EVOqTuDKNain1LBpsMC8lnJFirAoa8LzR6nqjv6fFiYUsCaEJiS/7fGZc" +
            "Elr+54rmGT827yFZtF/hbxucXUHzAQAA//8DAFBLAwQUAAYACAAAACEAN4zfm4QPAAC8oQAAEQAAAHdvcmQvZG9jdW1lbnQueG1s7B1rbxvH8XuB/ocDP8WA" +
            "TN77oUQqSB4ZBGgLI05+wIk8SWz4AklZdj/JdhI7tRvZDto4DiwnaYF+KNDqadGSLQP+BXt/ob+kM7t75PHpE2VHOh5l60Tu3c7tzM7MzszO7n70u+uVsnDN" +
            "bTRLtepCQkqKCcGtFmrFUnVlIfH5Z/nLZkJotpxq0SnXqu5C4obbTPxu8be/+Wh9vlgrrFXcaksAENXm/Hq9sJBYbbXq86lUs7DqVpxmslIqNGrN2nIrWahV" +
            "UrXl5VLBTa3XGsWULEoi/VRv1ApuswnvyzrVa04zwcEVroeDVmw461AZAaqpwqrTaLnXuzCkUwPRUlbKHAQkTwAIMJSlQVDKqUHpKWzVACB1IkDQqgFI2mSQ" +
            "hiCnTwZJHoRkTAZJGYRkTgZpgJ0qgwxeq7tVuLlca1ScFnxtrKQqTuOLtfplAFx3WqWlUrnUugEwRd0H45SqX0zQIqjVgVBRiqeGYKQqtaJbVoo+lNpCYq1R" +
            "nef1L3fqY9PnWX3+x6/RCIM/q2Jz5UAxTzXcMtCiVm2uluodCa9MCg1urvpAro1D4lql7D+3XpdCisso9WQzUnYBhmk+p3+lzFo+HqIkhugRBNGpEaYJve/0" +
            "W1IBLuy+eCLSBIgrhVQgPgB5AIBecEMqfB+GyWGkCl0JRTilkKLhw2G9gnBKXcJKIfVYf2MCAJrFVnH1VFBkn64prOu0nFWn2WF0hOierlFaB9yNSoBG9ZWz" +
            "CcLHjdpavQutdDZon3TV2joaGKeAxQUqKOTNszXm6qpTB21XKcx/slKtNZylMrQIxEMADhdoD+AVGAX/0I/udVqOfS2gjkksgmW0VCvewL91uKfO152G8wkw" +
            "pSrZtpa3wcLCUhhXWrQ0I9rZvGlB6TxYYcVPFxKimFV1JaN0iq40aKEt2qLUKbTdZWet3Bp8/EqgiLbiSgP/NOtOAVCEh8olJLSsYkvol0/XEE9nrVVLpPDR" +
            "Qq2KzVtzyldZJVr6pwI8fs0pLyQapZXVFnu0wYA38lClCfedZqEEHPFZqeI2hT+668KntYpTxRetpqvN4XcKzcFiCrz5Z/+NsuqXZPEtgbIUb0OKI4p/B2mv" +
            "qxnJNGW7l/YG/5nR/t3QHr4OJduFauz6fGuRbJGfyTOBvCYnAnlCtslL7553l7TJK+8mOSIn8P22QNoCPHfs3fe+gpLdwB1vA6q8gC+vsCpCbjH4o9hPtLKy" +
            "aBmZGfvN2A/Z7xHwzjHZg+se5aV9zlwvBLjswK1t76Z3i7RDsJah5GVDkeQZa81YC1nrR3JAXgDvHFDl9Mq7hZpsFz7cho8btPgVPiIAyx17t8lzsuPdC8Fn" +
            "iqFoOuOUGZ/N+IxsJcnTpEAeehvAWXdRe4VgIslIK5aWnSmrczCBDSuj5g1zRvuZADMB9m6BAUyHhH1qbWzDMNGmI8ext8lt4320fMlBZ+A4CW2XyDnTVJRM" +
            "dsZuM3ZjHhf8E978m/zYwzz+5UK19R/gE26RHwabGg1aXx7iX/wVrT74DGVvjsOYe2lFM0RtFjCZiS8VicdJ8gDMvaekTR3VYzoehDP65LysGhkxAqx0Hkwz" +
            "mTEna5ak5TSkXsTFs+DC/cZMPs8on9+TbW8TTDUmmG1yEEIwNdW2RUOfAhNtxkTvhIlOMFJ0Qg7JDg0M7XNeagvgF9xEhwDY61XHR7iP/gAaFMyNOIEnD2l9" +
            "5k7sku0wdkbOkM28pc548L3y4HDiG6ohy2ooBSBbCvwLSWf+8GR0LlWLcGe51Gi2fk+fNUSrf3heqrVWT0tp12m20s2Ss5DIOuXSUqN0nipAsTQdSHDxMQAB" +
            "xsl12oHQF/WG23Qb19zEInmEEeQTckJF/cC7xYUe1YV3U+DB5h0aP9imWmOTvKT3vFtJehEUcU5RBPLPJPk+Kfzvy78LsmwJCFAQ5aQkJmVRNNBj3CI74zTQ" +
            "y1H65+DNsdCjgi5WNyyxa7Y5mbKmOHt3MbyPH3HuchOofl8gR73Uokrc+3YkCb17OO+EkZ4jUPsv8XkaGmLxH3zo9vwoMiIRVV0XLS2KRBzO2SN5BpHV02o2" +
            "w1Rb5Djm0TjELNUw7YiKws/jpVyUVVkzI4kaaMVxnZZXVCmdjonoZXOanI+m6KF0yZapy+PwE3VFBFssJp2pqBkjF011AxbKtCqcCXpSyiqSGk3tKhljEdOU" +
            "vJjH1we6lhdGEdvkOGQzGcNQo6lcJXGsokkrRg7DCz1RA1Y4db2YVvWMje58EFleGEVkZVFWYqZqMdrWDcLto6OGPt1o9/dglPuL0TzqTtN5+7t+MG8c/xiK" +
            "ZOmZmAxcEbYnyTPygDwlW+QRTsk9Jk/g9xfyAPOVfyR/I9/hHSj61zj0c1lVVjAbKChOvDCKNJkbO5bnJEvuV4y8MB7cntZ0TbMj2bMf0AyGDXLg3aT50d5f" +
            "QJcdYcaSQMNb+zTs+Bx+X/oJSpfiZNiNjc5JspKOZlxnAia3csDn0VRfqqiYY92RtA62TjRt1rEGumFJtqr1SWOEQ3ZzwojlHduYVUmXd8z1LToSyN44CmmS" +
            "pqZx7VCQQrwwkhQKgWwE8YpX7IT8hzwEExPNzR/ABn1A/hurXgXfrDeNsg/hTEaWxLgMuzlDlQP5AtyVYIVR5O2xCipvq4aW60OWF8aju9O6movqHChOc7/E" +
            "ZIMpVcuSQR6MQ82WJdPq515eOHWiGl28RgUHj8hunHo3Oa1SOnYimM/F9PZihCdoxg+nGcWOpls7SchZ1Oxcf+TpXOXz1ANob1JvH35Dffno5zoKdHXkEXmO" +
            "ky0YYbztfYPRSJbwyCZodskhuPPUz2cJ+AI5RNefhgKe9y6fpBMzftIen+TpS2iERze8Tbh1RA7mhk/kCAC47X09h74IvhVf9wrT+byvTs+EUeqkRSTH2HzQ" +
            "vmRGIGEcuZaz2DEQo8OASJ0jDJuzHUZOvE1kZbbfzWHvGhJMGaUrAnoZm3Mvkh4IfogVt8hOQCC2Mf+3E/tCQHc6b8cNTvoShgOS4d2Hd47rqUBi+cXnUlzr" +
            "2l1K+tZFmfOCKgIvqiDte8khMUQq58dJnOVIIsn3oLcwxRoIDlX2k5L2IRZ/R34iTwVJtBRV0UTRVDW4+5g8hUJaYpmWbNIA5DPyDIp0MKTgf6+llRq6XEAy" +
            "zLyczqUTPcsFZE1MS3mla3XSUV00dFnvFAWWC+SzRjrbu1yAF9EXMuK9k+UCb1uwcZ6LCE5veKzPr/kFuFNd2Q2yG/YXx6FXi0Usjj4WyZGKDmSlo6L8kYAn" +
            "RUBhG5PZA0qMDtdT5FdMSLNdtnwDTAgYLVGTDLM8Bug03URByw5sBhwMRzFInMQKpQe5Ypde6UI/Zs7u4cz7LtsG7xCNLByI4iRTi4p2Yc2U94Dtm59i1bei" +
            "GaJz/SKmEKasx4crhDlRnNoBYfGDOLE4+YV6hvve19weAmuJLR6EC9hJ8OUOW8LZZltD0ZxUGA7QU7wfL6NAwB3uyE6vgxTgjxhZBLEa4sWBrKWYdu5UMvgi" +
            "3cbodfLSQGRw1svTrsYCbh56NSyQ7N2lvvAr717MXRtpHLIzzybK2MqaFaJz/aKp9GwW52R1pgIHh4Gp6d84+XEj4nZbZB8n+AK+HDpwuEqGTdG1O+4e7u9C" +
            "n7jFtnrZ5V5gvGwi7uDFapQf1ILT3ceYOgf2fqxU/yJ4N7jflm/s0o8o5js8SwStXCjE423oRt+Dlm+87KOxpm9OUc28Nj3IaoMBjiC68TAGw5Ag0tiuz5cd" +
            "mkhBy93q5c+v9hNhJA3iZA/HK/w/3ELkRiHmBQzucDTl/R+z6Kdv8opiKMNIV3UlQvvlvN0wildvX+rFNzV8M+BMWtRsqe+AL83OwU++m8qH9Bm6EzNPmPUL" +
            "3+NOzNFM+Sux6yn61e9C7C/e3KEkjxo+dAgK5OyxrHXMmGc7jh+Q57wAxqbu+USYWUzv4hYgzGehO35gzhp1VmAUw6zwY38f3NdwD79TCLe9b71vsIIA7g++" +
            "EtUeO9QTC+l7juhGIoc0n3wbn+weWnOaRN4QwqbJUl4TFUw6n2jb819P2GYiFRWR8jZS3k1BFQ1RNiVwXiy4wo+kWAJ5yFge08p/IQ+FD+gGUluX5mD0x0qK" +
            "KIkSPG5gBVrJNOaw1mPyRBBVVZM1KAkziki6pubU/MSHvfx6jB2vRPDo8rWDJBphqEV3//ZRmex+3vrXNDB/09/db/hCMZqS27c+hx7d190anqd5T7RLv6Do" +
            "73l3/v6d6zHb7A4dSV/D+A9N4oP2C+alvIR34Gh+hAuMaISSD/K3WZYaz1YOrF9qY+AzsDbJX4fETi2hDfYXR1HK+cv1cKHTEdobbfYingRHawBAtqCJ1qAW" +
            "BVQLRlFf42G7mCtHF1pRJPFwRb9dm2xpoL9AgQdeKYZ4Xi8SC1f24cGqDP0DbAftqT2WkEcT83run1DrAw/8vZcUyBNmNtGu4KZV5wHaEdvAYGwpF7z2GMDc" +
            "YdsTHQ1WxGadcBMLl1e0wWq6N3qyJJKCGSZiskPx/wZ5kh6oTI1PvtZzTvDu4Gq6t8kbNXBfsLKgRYrsQLuf8lBnEQuvjELkfcktUxQPdlj4AV8Aiq/13wfd" +
            "iG/FRmA9CrxfBqhMtrFjt+lREfBWtKSxkWAmQ2HgVAjkUeQHLqcMZ795OLkwBFcuY+POm+CmPqcEF6ShwuKvo+0Rlo6UYrLODmg/qjz6F+ZSCCydZ9ji2QFR" +
            "5/R4TJ4FyDOE7lR3MMpvhzGJNCkn6aYWgYN3z2YSXRzrZ3g/qHk9k1PpxgqzfjjHfjAsWbJse2Lf97yP/HqvPfFOnNgLe8DcMzqMBGwuchAmWiJZGTEH9uaM" +
            "Y2LHMVveRr+lefEa2Ua7nFovx289KrF/a4qR+6lc7I75MITYKpKdzll5FK+Z2MZLbJ9Ql7HjUqN7gI5Ezx443LfgrjvdSIcKy2YYo1rRVTsn2REw5iLCW8PJ" +
            "jFSWZSmyx8DPRPhMttow8eQTa6PkOoTs6oqs2lKuz5yT0lnLyurdfa6ROpqp5mTccCYM/+RsxciwLIUZ/5w7/wydOx1l7CAmQzevvKDojdyIZuzM8NuyjCKF" +
            "rHDGH/I4SR4k8fwajL8h3VCPwKehGqTpFlpXetilVzVchfu0VDUUhW2mWF+5iniug2qRLJFmsKziBKWpmAzp+sofHATZqtUXEprOhrPSyioAMjWqKJZqrVat" +
            "gjdNqjfcZbgnyQYD5jpFt4HZJ/Tmcq3WCnxdWWvRr6KvXspIYE5FfIYWF2uFjxslTGdBnXSl1Cqsom1DK6V8rOnHpVrxBv0AVdYqbrW1+H8AAAD//wMAUEsD" +
            "BBQABgAIAAAAIQDftUy2CgEAAL8DAAAcAAgBd29yZC9fcmVscy9kb2N1bWVudC54bWwucmVscyCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKyTTUvEMBCG74L/Iczdpl11Edl0LyLsVSt4zabTD2ySkplV++8NK7vb" +
            "xaV46HHeMM/7BJLV+tt24hMDtd4pyJIUBDrjy9bVCt6K55sHEMTalbrzDhUMSLDOr69WL9hpjkvUtD2JSHGkoGHuH6Uk06DVlPgeXTypfLCa4xhq2WvzoWuU" +
            "izRdyjBmQH7GFJtSQdiUtyCKocf/sH1VtQafvNlZdHyhQhIyx5tRZOpQIys4JElkgbyssJhVgYcOxwL7eao+m7Pe7Ii9fY9tR4MkOaWyZbTZlM1yThuOu3gy" +
            "2Y+/4aTD/ZwOlXdc6G038jhGUxJ3c0p84fb1z9schQcRefbt8h8AAAD//wMAUEsDBBQABgAIAAAAIQBQeo3y+gYAAPwgAAAVAAAAd29yZC90aGVtZS90aGVt" +
            "ZTEueG1s7Flbixs3FH4v9D8M8+74NuNLiBPssZ3bbhKym5Q8am15RrFmZCR5N6YEQvLUl5ZCWvrQQNuXPpTShaY0lIb+he1vCCT08iN6pLE9I1tukmYDoewa" +
            "1rp85+jTOUdHxzNnzt2JqbOPuSAsabnlUyXXwcmADUkSttwbu/1Cw3WERMkQUZbgljvDwj139v33zqDTMsIxdkA+EadRy42knJwuFsUAhpE4xSY4gbkR4zGS" +
            "0OVhccjRAeiNabFSKtWKMSKJ6yQoBrVH3xz9dPTr0aFzdTQiA+yeXejvUfiXSKEGBpTvKO14IfT17/ePDo+eHj0+Ovz9HrSfwvcnWnY4LqsvMRMB5c4+oi0X" +
            "lh6yg118R7oORULCRMst6T+3ePZMcSlE5QbZnFxf/83l5gLDcUXL8XBvKeh5vldrL/VrAJXruF69V+vVlvo0AA0GsPOUi6mzXgm8OTYHSpsW3d16t1o28Dn9" +
            "1TV821cfA69BadNbw/f7QWbDHCht+mt4v9PsdE39GpQ2a2v4eqnd9eoGXoMiSpLxGrrk16rBYrdLyIjRC1Z40/f69cocnqGKuWhL5RP5qrEXo9uM90FAOxtJ" +
            "kjhyNsEjNAC5AFGyx4mzRcIIAnGCEiZguFQp9UtV+K8+nm5pD6PTGOWk06GBWBtS/Bwx4GQiW+4l0OrmIM+fPHl2//Gz+z8/e/Dg2f0f5muvy11ASZiX++vb" +
            "T/9+dM/588ev/nr4mR0v8vgX33/04pff/k29NGh9fvji8eHzLz7+47uHFnibo708fJfEWDhX8IFzncWwQcsCeI+/nsRuhEheop2EAiVIyVjQPRkZ6CszRJEF" +
            "18GmHW9ySB824PnpbYPwTsSnkliAl6PYAG4zRjuMW/d0Wa2Vt8I0Ce2L82kedx2hfdvawYqXe9MJnANiUxlE2KB5jYLLUYgTLB01x8YYW8RuEWLYdZsMOBNs" +
            "JJ1bxOkgYjXJLtkzoikTukBi8MvMRhD8bdhm+6bTYdSmvov3TSScDURtKjE1zHgeTSWKrYxRTPPILSQjG8mdGR8YBhcSPB1iypzeEAthk7nKZwbdy5Bm7G7f" +
            "prPYRHJJxjbkFmIsj+yycRCheGLlTJIoj70oxhCiyLnGpJUEM0+I6oMfULLR3TcJNtz98rN9A9KQPUDUzJTbjgRm5nmc0RHCNuVtHhspts2JNTo609AI7S2M" +
            "KTpAQ4ydGxdteDYxbJ6RvhRBVrmAbba5hMxYVf0EC+zoYsfiWCKMkN3BIdvAZ3u2knhmKIkR36T5ytgMmR5cdbE1XulgbKRSwtWhtZO4KmJjfxu1XouQEVaq" +
            "L+zxOuOG/17ljIHM7f8gg19bBhL7K9tmF1FjgSxgdhFUGbZ0CyKG+zMRdZy02NQqNzIPbeaG4krRE5PkpRXQSu3jv73aByqM518+smCPp96xA9+k0tmUTFbr" +
            "m0241aomYHxI3v2ipoumyTUM94gFelLTnNQ0//uaZtN5PqlkTiqZk0rGLvIWKpmseNGPhBYPfrSW+JWfAo0IpTtyRvGW0GWQgFww7MOg7mgly4dQkwia8+UN" +
            "XMiRbjucyQ+IjHYiNIFly3qFUMxVh8KZMAGFlB626lYTdBpvs2E6Wi4vnnuCAJLZOBRii3Eo22Q6WqtnD/iW6nUv1A9mFwSU7OuQyC1mkqhaSNQXgy8hoXd2" +
            "LCyaFhYNpX4jC/019wpcVg5ST9F9L2UE4QchPlR+SuUX3j12T28yprntimV7TcX1eDxtkMiFm0kiF4YRXCarw8fs62bmUoOeMsU6jXrjbfhaJZWV3EATs+cc" +
            "wJmr+qBmgCYtdwQ/oaAZT0CfUJkL0TBpuQM5N/R/ySwTLmQXiSiF6al0/zGRmDuUxBDreTfQJONWrtTVHt9Rcs3Su2c5/ZV3Mh6N8EBuGMm6MJcqsc6+IVh1" +
            "2BRI70TDA2ePTvl1BIby62VlwCERcmnNIeG54M6suJKu5kfReB+THVFEJxGa3yj5ZJ7CdXtJJ7cPzXR1V2Z/vpm9UDnpjW/dlwupiVzS3HCBqFvTnj/e3iWf" +
            "Y5XlfYNVmrpXc11zkes23RJvfiHkqGWLGdQUYwu1bNSkdowFQW65ZWhuuiOO+zZYjVp1QSzqTN1bexHO9m5D5Hehep1SKTRV+BXDUbB4ZZlmAj26yC53pDPl" +
            "pOV+WPLbXlDxg0Kp4fcKXtUrFRp+u1po+3613PPLpW6ncheMIqO47Kdr9+HHP53NX/Xr8bXX/fGi9D41YHGR6bf4RS2sX/eXK8br/vQtv7Or5l2HgGU+rFX6" +
            "zWqzUys0q+1+wet2GoVmUOsUurWg3u13A7/R7N91nX0N9trVwKv1GoVaOQgKXq2k6DeahbpXqbS9ervR89p357aGnS++F+bVvM7+AwAA//8DAFBLAwQUAAYA" +
            "CAAAACEA62kGXHkYAAAbegAAEQAAAHdvcmQvc2V0dGluZ3MueG1stF3bblzHlX0fYP5B4PMoqvuFEzmo6yRBPBlEDua5Rbakhkk20d20rATz77OKF1FW1g7s" +
            "BIEBq9mrT506u/Z976rz69/8eH314oft4bjb37w+079SZy+2Nxf7y93N+9dnf/5uvkxnL46nzc3l5mp/s3199ml7PPvNN//+b7/+eH7cnk742fEFhrg5nl9f" +
            "vD77cDrdnr96dbz4sL3eHH+1v93eAHy3P1xvTvjz8P7V9ebw/d3ty4v99e3mtHu7u9qdPr0ySoWzx2H2r8/uDjfnj0O8vN5dHPbH/bvTuuR8/+7d7mL7+M/T" +
            "FYefc9+HS/r+4u56e3O6v+Orw/YKc9jfHD/sbo9Po13/o6MB/PA0yA9/7yF+uL56+t1HrX7G437cHy4/X/FzprcuuD3sL7bHIxbo+uppgrub5xu7vxno871/" +
            "hXs/PuL9ULhcq/tPX87c/7IBzN8MEC62P/6yMdLjGK9w5Zfj7C5/2Tjh8zi7Z8Lq8I9N5osBjpenyw+/aBTzRNdX69rNafNhc/zMRWvE7S+blP883KfrZxod" +
            "r34O1zxAf9i9PWwODzL5yDLXF+e/e3+zP2zeXmE6YJ0XWP0X97Nb/wcR1z/3H7c/3n+/6HD2DXTEX/b76xcfz2+3hwsIChSMUWevFnC5fbe5uzp9t3n75rS/" +
            "xU9+2GCSUaUH+OLD5rC5OG0Pb243F+Dhtr85HfZXT7+73P/3/tSgQw5g8ccr7jXK86c3D9oJV9xsrjHtn2icb/eXUB8fz+8Ou59P33XB/d21//KWX99oD216" +
            "2F1uv1vkenP6dLWdmPyb3V+25eby93fH0w4j3uudf2IGf28C25t15z9igb/7dLud283pDmT6F93sfiXm1e72293hsD/87uYS6/wvu9nu3bvtATfYbU7bb8E+" +
            "u8P+4z2df7vdXMKI/Yvue3fc/i9+DPmy34Etv6/702l//dtPtx9A639uJe/F4dWX7AtTfHl8+vCn/f70+acqm+jaowAt9BlRysQeBKSryRGbk4CEXDNHYk5N" +
            "QNqwFNGuGQHxbVSKGNVz5Ijukz+pccY/6o+vkaAVH82a4QxHnB6FI8FkPmunqub3cdCnnAbOaM/v44zxfLWdKUmYQXTlkZe/QrwP5lFxfY0EZYVrYs4S0hzn" +
            "naBHGhyxKnDeCVFHTp0QW+b3icZmvnLRFO844pzj94nRTk6dpKYwWrKx8rmlkB3n0ayGQLds6uR8kC3WmyPed07R7Lvlq5BDNMLconZ8tXPMwtyKap1fU/QQ" +
            "ZLuYWPgqFNc15/jiRhAQ7wVdVcLsfOWqcp7zTgVn8yetuk1O0Wqd52taQ3ecq2r0ms+gKWP4aE3NySnajJFGM3lwDmneZU635kPitG6+CRzfohc4voN9OEW7" +
            "z4XPugdduCbvwRdO0Q5Nwe8zYJo4jw4dQueISY0/z3B9cD6YNkVOt+lhTgTEVU6DGfugtkQr6z1dOQ3DMAQkjMgRbXWlFAUyu3CNq5wPtA7DURpoY0sXEBc7" +
            "1QcadtvS9QFiIpUFbZXXdLXxoI77IUBG4rS2NnDNByR1yqPaOheFGcQ0KR9oBzeNX+O0oKuATEtlWztbuDRq5waXH9ylZL4+LszIqeNi4vZUe5MSn4G3k/uj" +
            "2ns7OHWC8oXKnA56Rk6DYOF2Csi0fAbBJcM5PsBJ4usTIY98BkA8p0EMRrhPDF7gqgSbzmmddIv8eZKx3FvX8EO4LdHZNYFDcnCCpijw1zl1gDi+ckWNLiDO" +
            "ClJf/KicQ0oQvABdQ4p8tApdxSnaYNC41DcoEb4+zY7CV7u5zL1b3ULhNgsrCjYVkCjokK7bU3rla8TkwqnTfXScD3poliNDZcdnPXQ2nN9G1NwC6qk1j3L0" +
            "9IH7/pD5YujcjLKRc69REdERRSDAhd4HyNACAuNIuddAlTtKA2NU4RElkN4ojxo4dpPqXiCDWzMgM1FpNLDBSZiBM5nTzTiv+ZMaGFphbt5yfW1gtxuftYWH" +
            "T+XHWHhPVJMD6dxiGAuXjz+pjV1AnK6N0wB+A+dr42yIwjWwmpSvjTc68OfxFoTjiPeDUyeoyK2ZCc5E4RrnI59B1L1S/WaiL4nqEBOjYOdMgqwKiG08h2KS" +
            "a8JqJz8EWqeYg4QUwzkkW/zHkaANpwH0OPddgLjOZ51jinwGxXRus0yxkC2O+DA5V5WgB59biY7H9UBa43NrbggaqXklaBdEhzzHZbpygd+nG8etpul28ujQ" +
            "dG94/GN6UNxzMD1Gnk80QzlBHwzleZRjhlaR67ehE8+/ARkCvw0n5F3MCCVxvh4x8dytmcuD44gJia/PDLly6swo2TnEp1zH28XZdOUWwjWFVTFqSh2ro7Cm" +
            "QAK3P9boXKl2sTBNXLaBFE2pY6H+ed7Smqgq5R1r7eA+n7VOBbradsWNVLaBVOF5RKtpETVy2wgkCrN2EWqZIojaKr+PD0Ie1gaVA6d1MJ1nmGywdVCOt8EV" +
            "rimA9CbcJ1ThmmgT9/ls0krxGSSrhVknb7keBVIU594UtECdhOiQzzqFwTWSzcrz/AEWx3KvBogkC9mGwZ80+8DzOzYH1fmTQq54TGuL7by+YBEDDs5VxU/u" +
            "I9kSDc+T22oMz2AAqZnTuq4yj4BUHrXZCjnlfFB9D1yCGxworl0QHVpOtxYar0nYrgTf33btefRhu3Wdcwiiw8LXtEfHq3BAmuL3GVL0AWRy79YiQueRqx2+" +
            "Cxp2Qi1y3pmm8rjeTqd5jhhILsI1vkozgJzSlXMKskCfx61ImNpTh5iFa38HPcY9fCCFZ78c4mBLZ+20dty/hn9SEx/N6MLjRofA0UpI4Z6Qs2BtKgsIsyav" +
            "UTqocp4PAdK5HoU7bHndDEjgtRznTMt8TR2sMJUSIH0I1zjD4wUgXlgF56qnWgzIEDjERcdzgy7oxqvlcNUtj/RcgEnncwux83wvXJrA8/EuakSbHPGeawog" +
            "kVdsXIRO5LNORg1+TTJWWJ9kHa+fuuQdz0o5xICVPyniOe6puuICz7oDiTyH76ry3AYD6YKUVO25zwekcs/BVZCAz7opz/spELsHbpmAFIF7mxW0skOsya0M" +
            "kDz5mrbolXBNzDw3CGRUrhO7qtwPWZVq3u8CBGaGIsNMxUeDF5/5NWI85+aK3znihY6be0S4JioeHXqlGpcFD6+Pe+tYgsp5FIyjeXTotRLW1GvvCtVvXofZ" +
            "+NwMVAJ/HgNzRjnEr1QjpQ6QzCMjb3znHoo3QXPvFoirfNZWea5HvcVgAmI8jzGARIGi1iReTfLWFc9pjaCA13+8s45HoR5RqKYa1mN1BK7yStBI3sP35uvj" +
            "/Uj8Pj5knkPxCIN51OaDHorPLRih2ueDz1zHA5F4NMTBJdhHZXj/gY8m8nyIj4iD+colqFG+2kkL0ZRPK1fPEWt4rxQQL8h2gu4T7uMzt86Lntzz9tkoxbVL" +
            "dkI+0ecoREaI0BWvDPkCJ5bfp0B+OK1h5Xg2wleVJl+FamfmNKhuOK5dKvQ156oasxdGizNwHoXd5pbWt1gFpKvMK6u+a8F38d15QRa619xPROQceRzsx+qa" +
            "44hNAu8g1uR+oh+gG0em0jxq81OXwWcNS8sztGF1wFGpDwoiR+UnLItK6bY6W3nmFG534bnboK3jvn/QQfFKChAh+wXEcsmCsgydrk8w0fEuqtUkyqsiwWrN" +
            "NUWwtvFYM9ioeVY3rEiPP48znteMEBgpniUILgi9KwEhC694wlH2vLcISPGcQ2DoLJUsIIPn34IPjVuM4OPg1hmBHsbjSBC89RCi4plTIJ37yiHqwbubQjSe" +
            "d6yF6JR0zSKPgEzeZxdiMNwLCElHgauSTQKPJhcFWUje8ApUSMHwbGvIpmf+PHmlSjjiphVGi9HxJy0qc48rFCf0NAJpAg0Qnwau+aqLgkaqIfGO9tBUtZxu" +
            "TU1eTQrNNV47DM0H7tUA6QJFGxw4PoNuDPc6Q3eD+2+hRxh7AUk8LgkDaoTL3DBwXwQk897jMKyklYdTPC8GpHGLHkZo0gyi4T3bQDzfixCmqYKczuAr12Iz" +
            "CpW7qHTlFmO517xOC2TyjumorOLZfdgYwbeMGn6ncI0fvAMiwjDxnqxo4OTTVQAyee4JajTxvHJczR6Ur4F03kccEYXyOka0xnLLBMQLc7Ou8y7EaCWZizYW" +
            "HiFHZxPfHQVntHNNAT5UXCsDGTyuj95GXv8BkvjOhuid5rm0GHTjegeGSegAj3B4eL0eiDN8tQM8eb4KYB7uxwPpvM8BlqRMqvli9EKnX4zBcQ8yJpV57ikm" +
            "M6RrrBBnxeQ0zyvHJHnRQAaPkGNWQu4WiKD5YtbCvg8gg3uDEbae19pitoPb4Fic15x3YJ15TSIW2CZ+zdr7wleuKiH7Fat1gctCdVOQkgrdx+fWvOIVm9h1" +
            "5XsEYjda0DvdOp5tjX1lHSgyELnylRtGiLcjVof78XHqPLnMTSvssATSBbrNqHl+B4iQcU6rQ57yTlKmcG89rQ4ROoOkvJDnAyJ0OybIFn/StHaM0udJ2nRu" +
            "mZKW+gaTDoHXPuDGB54FTUYVgW5GD76jKhkXudQna5TwpBZ+ooQMXk1K1hpec012mRmKOOt5LJOc87zXHcGH455dik4La5qM5X0oKVnNLVNC/MNrLECESn7K" +
            "anCdCGRyS5uyV1xfpxwS749PiIy4pwpzrnhEmaru3FtP1WZeo0wVgT1f06Yyz2imZgK3zqn5wXOQqSvN65qp28J92NS90PuVBjQMp+gwhe+WTMN13jOXhhf2" +
            "xqYRG88rp2khkBxxkWv/tPbcUVpn5Sz3LYFUXpXPKii+ey1DwXLvNuvYuLeeoXd4p2w2NvCcQzZO0C7ZeCEvlsVu+7yYinIIkMIjo3X6AK/tyucSZGsb9+yy" +
            "9YVXHoA0LsFAhC6QvJxBym/ZKcX7ojM8ZWF9HKSe09pZYd8UmNpzq5ld1Jx7gSSBqzyiUM4HPnheBwTSJ+fesLo0OeIG95Ey4gXex5VDCJNTFMEmr70DiTxr" +
            "mKMPfIcYkMHrJTktVuAIAiNO0eQSj39yiiXzGcD35x2SQCzP7uesleZSkqXaYc5W8VwnkMT3sOfsEo8OgQy+Ry2XkHg3EJDM+0OAVF4/zRWRPZ9b04bbrNys" +
            "5Tu7c1s1aY4gmhHuEzuvjuXuFN/dnjs0Ob9PX5k+CRG4d6jCc9F5wFUVrgmSPgBS+JNOGwUdP53Qk5UndLyMUI4v6zQDSjcgs1PeKfC3eM21rL2plEeBCFX5" +
            "gqiAextFw5Wn1CmgJ69jAHHC82B5eJataFcnldMCAeYyV4zVXI8W4xPP9xaoRM47xWlhByyQyrs3izOV+2LFSb3uQCrvUy0O8ZwwWqg8R1xcNJx7gUQeOxev" +
            "Is8nFq8Lr1UXbws/66h413k39zokhMfOQAavdBVYWp6nABK5bwmk807MEpThWqxA6LndLsFO3mdXYDO5N1gC/AO+piFqQX4WwNc06sJjphKd4rEMkCasT/SW" +
            "+wclxilISVJCr0dJLvDO3wK3huveAsXD98/B6VU8kigQBm4xgDSefysV5pHrt2oU9xNL9UGQkuo7z2AUWFpeSVkIz+EDqdwLKM1F3mdXurLcSyvdd94rVYZy" +
            "fMdBgW0UaDD05N0mZaw9vRyxVZjbcIpX4YAE3lNSRhR8ZQhp1Zw60xVh1tN3wZrNmHkuuirjecRS4T/yOlNVeFK62lWF0ugMgAg51bpOeaP8VrXOXLKqto73" +
            "71QdPI+qIQqWVyuARK7JgRQevVejhX7Yapzh+g1I4F0TQCr30qrxlse0QBz30oB07otVExrPeVerK9+BVG3IXH6qM4rHZkAszwUA6dyiAxGyunU1XHLqwGpy" +
            "D6UiCuW2pK4dfPxJAziRI7DAXMPWCNvI75O05n1PNRnh9A4gk2f3a7KF26yaQiqconmVtDiiLc9GAOk8F1Czm9zWV4TvXL/VohzvLarFTN4zVwtiTc5VJQq5" +
            "QSDCPv4KueIWvVYrVKBqjZFHehV2ju9Ur81YgQbNOOE+HQEIn3Vfyyogg2eyajdwfCkydOU923U44ayJOpXm1Rcgwpl+dVrB762wP7xO21bLJaVBU67wPVBA" +
            "Ks9FNwzGeaeBC7h/AMTzWLNpkIA+KZDxdDDv1wg8VUpRIIb31TTtK+9XbgYOJOVEsJvnEQsQoQ+/Gd+4DQYiVP+BCLu9FsLziQ2MyPVbszpyO9esq7z631ad" +
            "iXOIs5NXUhocLl7/aR5Wkz9p0JH3srXgcubcG0LkGraFWLh+a9F67kG2iNH4ysWQ+NlALRkv0C3ZwbPULUXLfcuWreKdFi1L1f+GKIdXeVrRgXcLt2KFrBQQ" +
            "x2tgQCLvem0Fy81XAYE998WACN5TKyFLo0VhNx6Qzs/FaRUBIp911VmgdTWWZ7KATMWpU23ieaRWXRNoUH3iNXEglZ8L2qBeuAVssMDcs2tddX5iG9SR5rWc" +
            "1tduUo7EwvtUG6I2HrEshPecNkRt3MNvwxvucTX4Qbwq0qYRdj4CEfK9bYbGK+xdqcSj3b7KTHQGfZ2eQZ+0Kx+5TuyrL4xyfNdacYveNdQon5v2Qgf42qjB" +
            "M8HdaMdtMJDE6yV9na5C+Q1I5JnGDgto6GoDmbxnAW5VdfyadXASpw6MM/eRug2dn/4JZPJdmRCFwbN53YXB96R0rzTPIwFJPPPTvTZcKwOpXE474i8eSXQf" +
            "NNdV3cfA+yB7UJprCiCF650e4JJzGgQb+bmtPXgtyE/wQvdmD8HwWKZHOH1UUwBJPL/T4zo9UEA611U9hspPj+oxRt7N0JNu/JSDnsBWnENS1IIOSRGhFkWy" +
            "ydwb7NkJJ5j07PvTOzi+QoptvD7XSxR263fYOX5GVG9GCfLTvOUdXr2Fwfdid0Rggpx2K5y33rsXau+9B8EP6bBzPCPTEc/xbHgfJgscP0Ll3lOfNgoaaSLQ" +
            "o7IwlA98Zx2QyNdnaK25dUYoZfnuKCBCZ88wwfDocMC15DZ4WOs4dYZ1icvPsL7w9RnrLBCOOFV5bnCs1CBdUyj4zCPK4deR0RxBdMhp4GG0+H28bbzTHIiw" +
            "R2B46dxwIIPvXxg+CHmx4aPl9mdAx/OTIkdA6Ez1KJDAPYcRrOJdyUA6t2YjBC9wCFQi10gjmsL38Y8ondE+ohPOiBrJF747d6QgnBcwUhyJyvZYHR181kVP" +
            "3gc5ChxIzonrNCy+PnWdGcoRl3nfBpAirFzTnWdBB6IP3kMLRMiCjrYOVKVIV4bvHRsdfi9fha6zwNdwO/meISCFV4ZGt5HHzqN74dTuAbeGn+IyehTe0DGG" +
            "CjwXMIYX6jJjrjctCEjip/CNqZuw2lMPgRMnnH/OO0B4tXxMOxWX0+mnoEPA7zy7jwft3I+fynkeo8+1Y4au6VwvOaCzntpWvrN7GhP4qZwTgSu3jdM6zTNM" +
            "QIR9H9MG4Qz9icF4/8F03vBa9fRYHn6NN4XXg6d3lncQAYmdPw9cCq4TJ5wnXm+c0OTCyoUwuX4D4wjnRszoMt8ZNJfI8dVOCF353BJCVyolQDz3E2eyitez" +
            "Jvx4YbVTFM4Xm9lUbhuBTC6nMzvBU53ZB+49zRyMQLeCqIBfUxA0SUjmOnGuMimX4Go79/1nDZVHbbNJJ0nPpqvAVc1UngGcLRZhfbrSPA5eCN/zMLvR3P7M" +
            "7gS7DSRyPTpXKE51PBDhlIOJYJNXCOewldfA5lhHwwmIsEdgwuUSaDBd4XnLCSXCfdg5Y3uw9a8eoOM3v74+X69R/Z/D06f16sQX1w9XtM3128Nu8+Lb9aLV" +
            "V+sXbw/f193NE/52+25/2H6JvLl7+wS+fPkAHK83V1fzsLl4Au4X7vr8cne87dt395+vvt0c3j+P+/iLA/32cvvu95/HWu+43B7+67C/u31APx42tw+vRHz6" +
            "iXYPB8pcn+9uTn/YXT99f7x7++bpqpvN4dMX0N3N5R9/ONzT6Zk8H89PH7bX96+W/MPm/lWF97893L38058fiH1xdXizXkO4/XZze/vwNsO37/Xrs6vd+w8n" +
            "vV5AeMJfl5vD9/d/vH1vHjFzj5kH7P6PzcV6Mvz68cPzd+bpuy9+Z5++s8/fuafv3PN3/uk7//xdePourO8+fLrdHq52N9+/Pvv8cX3/bn91tf+4vfztM/43" +
            "Xz0Q4fhhc7vtD+8aBXvtH754fPno8cUP59sfT6/Ptpe709mL4+3u8nrzI9bocRPi46+vNp/2d6ef/HZh68e3Px1hvcD18VWSr35y8T2LfzWX9Q7Uix3Y8c2n" +
            "67fPrzb9j4eJX+2Opzfb281hc9ofnrD/vMe0O7/cX/wOkoRP99+bddhzeNyNpP1n2D/Af13H7Vfd2ssZ8nwJcZwv05j1JcxR7dM3b237v0dBfHqr8zf/DwAA" +
            "//8DAFBLAwQUAAYACAAAACEAt/GuJK8AAAAOAQAAEwAoAGN1c3RvbVhtbC9pdGVtMS54bWwgoiQAKKAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAArI/BCsIwEER/JezdpnoQKW2lIJ5EhCp48JKm2zaQ7JYkiv69QcQv8Dhv4A1Tbp/Oigf6YJgqWGY5CCTNvaGxgst5v9iACFFRrywTVkAM27rsipbvXmMQ" +
            "LVrUEfs2vmyqb82pya7tAcQHHJVLMDEQaYdC0VUwxTgXUgY9oVMh4xkpdQN7p2KKfpQ8DEbjjvXdIUW5yvO17ExnDY9ezdPrK/uLqi7l70z9BgAA//8DAFBL" +
            "AwQUAAYACAAAACEADZre0uEAAABVAQAAGAAoAGN1c3RvbVhtbC9pdGVtUHJvcHMxLnhtbCCiJAAooCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AACckMFKxDAQhu+C7xDmnk23utu6NF26xsJeRcFrNp22gSYpSSqK+O6meFqPnoZvhpnvZ6rjh5nIO/qgneWw3WRA0CrXaTtweH1paQkkRGk7OTmLHKyDY317" +
            "U3Xh0MkoQ3QezxENSQ2d6llw+Lrbt+KpLHY0z7Mdvd8+lrRpxImKoilPD20h2rz4BpLUNp0JHMYY5wNjQY1oZNi4GW0a9s4bGRP6gbm+1wqFU4tBG1meZXum" +
            "lqQ3b2aCes3zu/2MfbjGNdri9X8tF32ZtBu8nMdPYHXF/qhWvnpF/QMAAP//AwBQSwMEFAAGAAgAAAAhAOP3lSJADAAA2XUAAA8AAAB3b3JkL3N0eWxlcy54" +
            "bWzEnc1y27oVx/ed6TtwtGoXjiR/yEnmOnccO6k9N058I7tZQyRkoSYJlaRiu7vbbR+g+z5BpzOd6WSmfQbnjQqAlAT5EBQPeOpuEuvj/ADij/8hDj/EH368" +
            "T+LgK89yIdOj3vDFoBfwNJSRSG+OetdX73de9oK8YGnEYpnyo94Dz3s/vvn1r364e50XDzHPAwVI89dJeNSbFcX8db+fhzOesPyFnPNUfTiVWcIK9TK76Scs" +
            "u13Md0KZzFkhJiIWxUN/dzAY9SpM1oYip1MR8lMZLhKeFia+n/FYEWWaz8Q8X9Lu2tDuZBbNMxnyPFcbncQlL2EiXWGG+wCUiDCTuZwWL9TGVD0yKBU+HJi/" +
            "kngNOMABdgFgFPJ7HONlxeirSJsjIhxntOKIyOL4dcYC5FERzVCU3eW49nUsK9iM5TObyHGdOljhHhI9Rkn4+vwmlRmbxIqkVA+UcIEB63/V9uv/zJ/83ryv" +
            "N6H3RnkhkuEpn7JFXOT6ZXaZVS+rV+a/9zIt8uDuNctDIa5UB1UriVANnh2nueipTzjLi+NcsNoPZ/qP2k/CvLDefisi0evrFvM/qQ+/sviot7u7fOdE92Dj" +
            "vZilN8v3ssXO52u7J0c9nu5cj/VbE8U96rFsZ3ysA/vVhpX/W5s7f/rKNDxnoTDtsGnBlc2Ho4GGxkJnld2DV8sXnxd68NmikFUjBlD+v8L2wYgr96tcMC5T" +
            "kvqUTz/I8JZH40J9cNQzbak3r88vMyEzlXaOeq9Mm+rNMU/EmYginlpfTGci4l9mPL3OebR+/+f3JnVUb4Rykaq/9w5HZhbEefTuPuRznYjUpynTmnzUAbH+" +
            "9kKsGzfhf1zChpUSdfEzznQ2DoZPEab7KMSujsitra1nLp5su/kWqqG952po/7kaOniuhkbP1dDhczX08rkaMpj/ZUMijVTiN9+HzQDqNo7DjWiOw2xojsNL" +
            "aI7DKmiOwwlojmOiozmOeYzmOKYpglPI0DULrcm+55jtzdzt+wg/7vZdgh93+x7Aj7s94ftxt+d3P+72dO7H3Z69/bjbkzWeWy61gnNls7To7LKplEUqCx4U" +
            "/L47jaWKZUpUGp7e6fGMZCMJMGVmq3bEnWkhM6+3zxBjUv/9eaErvUBOg6m4WWQ879xxnn7lsZzzgEWR4hECM14sMseI+MzpjE95xtOQU05sOqiuBIN0kUwI" +
            "5uac3ZCxeBoRD9+SSJIUVhNa1c8zbRJBMKkTFmaye9ckI8sPH0Tefaw0JHi7iGNOxPpIM8UMq3ttYDDdSwOD6V4ZGEz3wsDSjGqIKhrRSFU0ogGraETjVs5P" +
            "qnGraETjVtGIxq2idR+3K1HEJsXbq45h+2N3J7HUJxU692MsblKmFgDddzfVMdPgkmXsJmPzWaCPStdj7W3GtvNWRg/BFcU+bUWiWtebKXKitlqki+4DukGj" +
            "MteKR2SvFY/IYCted4tdqGWyXqCd0dQz48WkqDWtIbUy7ZjFi3JB291trOg+w9YGeC+ynMwG9ViCGfxRL2e1nBSZb93L7h1bs7rb6mlWIu1ehSToZSzDW5o0" +
            "fPYw55kqy247k97LOJZ3PKIjjotMlnPNtvyukaSV5d8l8xnLhamVNhDtd/XLyxGCCzbvvEGXMRMpjW7vdhIm4oBuBXF2dfEhuJJzXWbqgaEBvpVFIRMyZnUk" +
            "8Ddf+OS3NB08VkVw+kC0tcdEh4cM7EQQ7GRKkoyISGqZKVJBsg81vJ/4w0SyLKKhXWa8vAKo4ETEMUvm5aKDwFsqL96p/EOwGjK837NM6ONCVKa6IoFZhw3z" +
            "xeQPPOye6j7KgOTI0KdFYY4/mqWuiabDdV8mbOC6LxGMmmr3oOcvwcZu4Lpv7AaOamNPYpbnwnkK1ZtHtblLHvX2di/+Kp6MZTZdxHQDuASSjeASSDaEMl4k" +
            "aU65xYZHuMGGR729hFPG8AgOyRne7zIRkYlhYFRKGBiVDAZGpYGBkQrQ/QodC9b9Mh0L1v1anRJGtASwYFTzjHT3T3SWx4JRzTMDo5pnBkY1zwyMap7tnQZ8" +
            "OlWLYLpdjIWkmnMWkm5HkxY8mcuMZQ9EyHcxv2EEB0hL2mUmp/rWEJmWF3ETIPUx6phwsV3iqET+widkXdMsyn4RHBFlcSwl0bG19Q7HRG5eu7YtzNzJ0bkL" +
            "lzEL+UzGEc8c2+SOVfXyuLwt42n3TTdaHfb8IG5mRTCerY7225jRYGvksmDfCNveYN2Yj5b3s9SFXfBILJJlR+HNFKO99sFmRm8E728PXq8kNiIPWkbCNkfb" +
            "I9er5I3Iw5aRsM2XLSONTzcim/xwyrLb2olw2DR/VjWeY/IdNs2iVXBts00TaRVZNwUPm2bRhlWC4zDUZwugOu08445vZx53PMZFbgrGTm5Ka1+5EU0G+8y/" +
            "Cr1nxyRN097q6gmQ980iulXm/Hkhy+P2Gyec2t/Uda4WTmnOg1rOXvsTVxtZxj2OrdONG9E677gRrROQG9EqEznDUSnJTWmdm9yI1knKjUBnK7hHwGUrGI/L" +
            "VjDeJ1tBik+26rAKcCNaLwfcCLRRIQJt1A4rBTcCZVQQ7mVUSEEbFSLQRoUItFHhAgxnVBiPMyqM9zEqpPgYFVLQRoUItFEhAm1UiEAbFSLQRvVc2zvDvYwK" +
            "KWijQgTaqBCBNqpZL3YwKozHGRXG+xgVUnyMCiloo0IE2qgQgTYqRKCNChFoo0IEyqgg3MuokII2KkSgjQoRaKOWtxr6GxXG44wK432MCik+RoUUtFEhAm1U" +
            "iEAbFSLQRoUItFEhAmVUEO5lVEhBGxUi0EaFCLRRzcnCDkaF8Tijwngfo0KKj1EhBW1UiEAbFSLQRoUItFEhAm1UiEAZFYR7GRVS0EaFCLRRIaJpflanKF2X" +
            "2Q/xRz2dV+y3P3VVdeqzfSu3jdprj1r2ys1qfy/CWylvg9obD/dMvdEOIiaxkOYQteO0us01l0SgTnx+Omm+w8emd/zRpepeCHPOFMD320aCYyr7TVPejgRF" +
            "3n7TTLcjwapzvyn72pFgN7jflHSNL5cXpajdEQhuSjNW8NAR3pStrXA4xE052gqEI9yUma1AOMBN+dgKPAh0cn4afdBynEar60sBoWk6WoRDN6FpWkKtlukY" +
            "GqOtaG5CW/XchLYyugkoPZ0YvLBuFFphN8pPamgzrNT+RnUTsFJDgpfUAOMvNUR5Sw1RflLDxIiVGhKwUvsnZzfBS2qA8Zcaorylhig/qeGuDCs1JGClhgSs" +
            "1B13yE6Mv9QQ5S01RPlJDRd3WKkhASs1JGClhgQvqQHGX2qI8pYaovykBlUyWmpIwEoNCVipIcFLaoDxlxqivKWGqCapzVGUDalRClvhuEWYFYjbIVuBuORs" +
            "BXpUS1a0Z7VkETyrJajVUnNctWSL5ia0Vc9NaCujm4DS04nBC+tGoRV2o/ykxlVLdVL7G9VNwEqNq5acUuOqpUapcdVSo9S4asktNa5aqpMaVy3VSe2fnN0E" +
            "L6lx1VKj1LhqqVFqXLXklhpXLdVJjauW6qTGVUt1UnfcITsx/lLjqqVGqXHVkltqXLVUJzWuWqqTGlct1UmNq5acUuOqpUapcdVSo9S4asktNa5aqpMaVy3V" +
            "SY2rluqkxlVLTqlx1VKj1LhqqVFqXLV0oUIEwU9AjROWFQHd78WdsXxWsO4/TnidZjyX8VceBbSb+gG1lf27jcdfabZ5Np/6fqHGTP8CunW7UlT+AmwFNF88" +
            "VyRmnmClOxFUzwKrHlxl+lqdqS0bMzGwlXCmmgmrn61ytTIAzTh+kdY0u55py29XY7cemPJ7G8PS2MtCz+ymHg4dA1F6wtWvV5XJt3VMdWMSl49EU3+cp5EC" +
            "3FWPAys7GN2zEqU+P+FxfMHKb8u5+6sxnxblp8OB+UmCJ59Pyl/Xc8ZnJg07Af3NzpQvq8eyOYa5/L396voA11Dv1gy1uVCl6yi7+7VhhXVP9kBPNm4ULweR" +
            "KfgnbVXzheXQK6uu3qrO73eZHlku9JwwUYPB7ulosHdaxrqepGc/R29/9aL+OXqOhxGqvMNvJA+uz3W4edDg5lthbr0ut2T1bMFhtYe0ny1Yvmc9IrBNzggX" +
            "uZqlJomBqbIPBHr82+M/H799/+X7n4PHf3z/y+O/H//z/ZfHb4//Ch7/ql78/fFbvWzL3aitW3VFR2vd3CL9f4d3+Vf+5r8AAAD//wMAUEsDBBQABgAIAAAA" +
            "IQCoXBDXywEAAC0FAAAUAAAAd29yZC93ZWJTZXR0aW5ncy54bWyclF1v2yAUhu8n7T9Y3Dd2rCaKrCaVoqrTpO5DW7d7DDhGA44FJI7763fATuItvah7Yw4v" +
            "vI/PAR/f3R+1Sg7COglmTeazjCTCMODS7Nbk1/PjzYokzlPDqQIj1qQTjtxvPn64a4tWlD+F97jTJUgxrtBsTWrvmyJNHauFpm4GjTC4WIHV1OPU7lJN7Z99" +
            "c8NAN9TLUirpuzTPsiUZMPYtFKgqycQDsL0Wxkd/aoVCIhhXy8adaO1baC1Y3lhgwjmsR6uep6k0Z8z89gqkJbPgoPIzLGbIKKLQPs9ipNUFsJgGyK8ASyaO" +
            "0xirgZGic8yRfBpneeZIPuK8L5kRwHHP60mU/HSuafBST2vq6jFRTEtqccZ1OpyRZsXnnQFLS4UkvPUELy6J4PDE+sMQQ3GMeiiBbLAhuDy4YUzaIhzxPF/c" +
            "rnL8lhZxQwm8e4iLB6pwlaRBxX54EpU/qdlZ/SF39SvyMzTX4ha8B/2fjolsuQ2Rv3gM9jHBiXsJ+0LQUCaGmIECbD+699Aj1Cizac7yn4ymee248inW9FJ0" +
            "H57GeDHQeKnli3gEu7XQOmH7twnVfTO/vzzFGVUK2u9fP/W00U9u8xcAAP//AwBQSwMEFAAGAAgAAAAhAC6DJEwLAgAAhQcAABIAAAB3b3JkL2ZvbnRUYWJs" +
            "ZS54bWzck9uK2zAQhu8LfQej+41lxzk0rLOwaQILpRft7gMosmyLWpLRKKe370h20kAIrAttoQbb8j+azzO/PY9PR9VEe2FBGp2TZERJJDQ3hdRVTt5eNw9z" +
            "EoFjumCN0SInJwHkafnxw+NhURrtIMJ8DQvFc1I71y7iGHgtFIORaYXGYGmsYg4fbRUrZn/s2gduVMuc3MpGulOcUjolPca+h2LKUnLx2fCdEtqF/NiKBolG" +
            "Qy1bONMO76EdjC1aa7gAwJ5V0/EUk/qCSbIbkJLcGjClG2EzfUUBhekJDSvV/AJMhgHSG8CUi+MwxrxnxJh5zZHFMM70wpHFFef3irkCQOGKehAlPfsa+1zm" +
            "WM2gviaKYUVNLriT8h4pvniptLFs2yAJv3qEHy4KYH/F/v0tLMUx6L4FsuxHITosNFOYuWKN3FoZAi3TBkSCsT1rcoI9bOiE+l5SmtGxv5LYb+Q1syA8JGxc" +
            "rTq5ZEo2p7MKBwnQBVrpeH3W98xKX3UXAllhYAdbmpN1Rmm63mxIpyRIpqhks+deSbGo7vjUK+OLQr3CAyc8Jh2HB85lD74z7hy4ceJVKgHRV3GIvhnF9B1H" +
            "UjpFJyboh3dmPMgRG7iDHPH93zgym0/+iiPfRWVE9PZyx4rnYEHWn2jGP/g51tnsrPxRK/oxib7IqnZ3h8WPyH86LP0Clj8BAAD//wMAUEsDBBQABgAIAAAA" +
            "IQCIqFXXnAEAAA4DAAARAAgBZG9jUHJvcHMvY29yZS54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACcks1O3DAQgO+V+g6R71nbSYHIygaprTiBhESqIm6WPSxWEyeyXZa9wYlrj730HSraSggEz+B9I5zsJvwU" +
            "qVJvHs/nz+Px5NtndRWdgrGq0VNEJwRFoEUjlZ5N0adyJ85QZB3XkleNhilagEXbxds3uWiZaAzsm6YF4xTYKJi0ZaKdohPnWoaxFSdQczsJhA7J48bU3IXQ" +
            "zHDLxRc+A5wQsolrcFxyx3EnjNvRiNZKKUZl+9VUvUAKDBXUoJ3FdELxI+vA1PbVA33mCVkrt2jhVXRIjvSZVSM4n88n87RHQ/0UH+7tHvRPjZXueiUAFbkU" +
            "TBjgrjGF/+Gv/Y2/97f+j7/rVpH/1gc3ywv/M2z9Xp5H/rv/5e+X52H7yl8vL3P8RNG1u+LW7YWfOVYg3y/+2/q3qZMbOFXdDBRpT4zhcPG+UdqBLBKSbMQk" +
            "i5OkpO8Y3WKEHI3OAcrXv7AqH2QUusdWvR4yn9MPH8sd1Pk2Y7IV07QkGaPZyvfi/KOwXlf9T2MWk42Spiyhz42DoOiLfj7BxQMAAAD//wMAUEsDBBQABgAI" +
            "AAAAIQBeGTa6+gEAAOsDAAAQAAgBZG9jUHJvcHMvYXBwLnhtbCCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJxTy47TMBTdI/EPUfZT96XRTOV6hDpCgyi0UtOZtXFuWgvHtmxPNeVf+AeEhGDDP+STuE6mwQVWZBGd" +
            "+/DxuQ/Tm6daZQdwXho9z0eDYZ6BFqaUejfPt8Xri6s884HrkiujYZ4fwec37OULunbGggsSfIYU2s/zfQh2RogXe6i5H2BYY6QyruYBTbcjpqqkgFsjHmvQ" +
            "gYyHw0sCTwF0CeWF7QnzjnF2CP9LWhoR9fn74miRj9ECaqt4APY+nlSU9A5amMBVIWtgY3T3Bl3zHXg2oqQD9MG40rPJ9IqSDtLFnjsuAjaPja6vh5QkDvrK" +
            "WiUFD9hX9k4KZ7ypQrZqxWaRgJI0hWIBGxCPToYjQ6rUpEupo5RLSjqE2hzfOW73nk2jwN6iG8EVLLB2VnHlgZLfDnoHPM51zWUUeAizA4hgXOblJ5zsOM8+" +
            "cA+xY/P8wJ3kOuRdWme0WFkfHGs+N1+a781X/P9sfjTfKOlDLUxPpFhOY0M7cJ7YGq0cxOdCCxkU+FWFZYZ/6B6lulsNnepETqrsdMcfrAtTW66PbLl9u3qz" +
            "xFE+27H3H/3WFuY2bstzU8+dySI8yLDfWC5wSOPJZJquRBKiG/RCiTPup9Q76B0W4lS8AM/qHZSnnL8Dccnuu7eL2zEY4tdu1cmHq9E/KvYLAAD//wMAUEsD" +
            "BBQABgAIAAAAIQB0Pzl6wgAAACgBAAAeAAgBY3VzdG9tWG1sL19yZWxzL2l0ZW0xLnhtbC5yZWxzIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjM+xisMwDAbg/eDewWhvnNxQyhGnSyl0O0oOuhpHSUxjy1hqad++" +
            "5qYrdOgoif/7Ubu9hUVdMbOnaKCpalAYHQ0+TgZ++/1qA4rFxsEuFNHAHRm23edHe8TFSgnx7BOrokQ2MIukb63ZzRgsV5QwlstIOVgpY550su5sJ9Rfdb3W" +
            "+b8B3ZOpDoOBfBgaUP094Ts2jaN3uCN3CRjlRYV2FxYKp7D8ZCqNqrd5QjHgBcPfqqmKCbpr9dN/3QMAAP//AwBQSwECLQAUAAYACAAAACEAPlJI6HEBAACk" +
            "BQAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQAekRq37wAAAE4CAAALAAAAAAAAAAAAAAAAAKoDAABfcmVs" +
            "cy8ucmVsc1BLAQItABQABgAIAAAAIQA3jN+bhA8AALyhAAARAAAAAAAAAAAAAAAAAMoGAAB3b3JkL2RvY3VtZW50LnhtbFBLAQItABQABgAIAAAAIQDftUy2" +
            "CgEAAL8DAAAcAAAAAAAAAAAAAAAAAH0WAAB3b3JkL19yZWxzL2RvY3VtZW50LnhtbC5yZWxzUEsBAi0AFAAGAAgAAAAhAFB6jfL6BgAA/CAAABUAAAAAAAAA" +
            "AAAAAAAAyRgAAHdvcmQvdGhlbWUvdGhlbWUxLnhtbFBLAQItABQABgAIAAAAIQDraQZceRgAABt6AAARAAAAAAAAAAAAAAAAAPYfAAB3b3JkL3NldHRpbmdz" +
            "LnhtbFBLAQItABQABgAIAAAAIQC38a4krwAAAA4BAAATAAAAAAAAAAAAAAAAAJ44AABjdXN0b21YbWwvaXRlbTEueG1sUEsBAi0AFAAGAAgAAAAhAA2a3tLh" +
            "AAAAVQEAABgAAAAAAAAAAAAAAAAApjkAAGN1c3RvbVhtbC9pdGVtUHJvcHMxLnhtbFBLAQItABQABgAIAAAAIQDj95UiQAwAANl1AAAPAAAAAAAAAAAAAAAA" +
            "AOU6AAB3b3JkL3N0eWxlcy54bWxQSwECLQAUAAYACAAAACEAqFwQ18sBAAAtBQAAFAAAAAAAAAAAAAAAAABSRwAAd29yZC93ZWJTZXR0aW5ncy54bWxQSwEC" +
            "LQAUAAYACAAAACEALoMkTAsCAACFBwAAEgAAAAAAAAAAAAAAAABPSQAAd29yZC9mb250VGFibGUueG1sUEsBAi0AFAAGAAgAAAAhAIioVdecAQAADgMAABEA" +
            "AAAAAAAAAAAAAAAAiksAAGRvY1Byb3BzL2NvcmUueG1sUEsBAi0AFAAGAAgAAAAhAF4ZNrr6AQAA6wMAABAAAAAAAAAAAAAAAAAAXU4AAGRvY1Byb3BzL2Fw" +
            "cC54bWxQSwECLQAUAAYACAAAACEAdD85esIAAAAoAQAAHgAAAAAAAAAAAAAAAACNUQAAY3VzdG9tWG1sL19yZWxzL2l0ZW0xLnhtbC5yZWxzUEsFBgAAAAAO" +
            "AA4AlAMAAJNTAAAAAA==";

        public static void CreateDefaultTemplate(string targetPath)
        {
            string dir = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            byte[] bytes = Convert.FromBase64String(EmbeddedMasterTemplateB64);
            File.WriteAllBytes(targetPath, bytes);
        }
        #endregion

        static XmlDocument LoadDocXml(string path, out XmlNamespaceManager ns)
        {
            string xmlText;
            using (FileStream fs = File.OpenRead(path))
            using (var z = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = z.GetEntry("word/document.xml");
                using (Stream es = entry.Open())
                using (var ms = new MemoryStream())
                {
                    es.CopyTo(ms); ms.Position = 0;
                    using (var reader = new StreamReader(ms, Encoding.UTF8))
                        xmlText = reader.ReadToEnd();
                }
            }
            var doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.LoadXml(xmlText);
            ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            return doc;
        }

        static string ComposeSegment(string label, string raw)
        {
            string formatted; long rub; string kk;
            if (!TryParseAmount(raw, out formatted, out rub, out kk))
                throw new ArgumentException("Неверная сумма \"" + raw + "\" для \"" + label + "\"");
            string words = NumToRuWords(rub);
            return "сумму " + label + " в размере " + formatted + " руб. (" + words + " руб. " + kk + " коп.)";
        }

        // Генерация нового заявления на основе существующего шаблона
        public void GenerateApplication(string templatePath, string outPath, GenParams p)
        {
            // --- валидация ---
            string rawNum = (p.Num ?? "").Trim();
            if (rawNum.Length == 0)
                throw new ArgumentException("Укажите номер дела или заявления.");
            DateTime dt;
            if (!DateTime.TryParseExact(p.Date, "dd.MM.yyyy",
                System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dt))
                throw new ArgumentException("Дата должна быть в формате дд.мм.гггг (сейчас: \"" + p.Date + "\").");
            var fios = p.Debtors.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim().Replace("\r", " ").Replace("\n", " ").ToUpperInvariant()).ToList();
            if (fios.Count == 0) throw new ArgumentException("Укажите хотя бы одного должника (ФИО).");
            string addr = (p.Addr ?? "").Trim().Replace("\r", " ").Replace("\n", " ");
            if (addr.Length == 0) throw new ArgumentException("Укажите адрес объекта недвижимости.");
            string rep = (p.Rep ?? "").Trim().Replace("\r", " ").Replace("\n", " ");

            var segs = new List<string>();
            if (p.UseOsn) segs.Add(ComposeSegment("основного долга", p.Osn));
            if (p.UseGos) segs.Add(ComposeSegment("госпошлины", p.Gos));
            if (p.UseSud) segs.Add(ComposeSegment("судебных расходов", p.Sud));
            if (segs.Count == 0) throw new ArgumentException("Отметьте хотя бы один вид взыскания.");

            string debtorsClause = string.Join(", ", fios.ToArray()) + (rep.Length > 0 ? ", " + rep : "");
            string sumsText = "Взыскатель просит взыскать" +
                (p.Solidary ? " в солидарном порядке" : "") + " " +
                string.Join(", ", segs.ToArray()) + ".";

            // --- загрузка шаблона ---
            XmlNamespaceManager ns;
            XmlDocument doc = LoadDocXml(templatePath, out ns);

            XmlElement mainP = null, sumsP = null;
            List<XmlNode> mainTs = null, sumsTs = null;
            string mainS = null, sumsS = null;
            foreach (XmlNode pNode in doc.SelectNodes("//w:body//w:p", ns))
            {
                var pe = (XmlElement)pNode;
                var ts = pe.SelectNodes(".//w:t", ns).Cast<XmlNode>().ToList();
                string s = ParaText(ts);
                if (mainP == null && (s.Contains("ВС№") || s.Contains("229") || s.Contains("исполнительном производстве") || s.Contains("прошу принять")) && s.Contains("адрес объекта недвижимости"))
                { mainP = pe; mainTs = ts; mainS = s; }
                if (sumsP == null && s.StartsWith("Взыскатель просит"))
                { sumsP = pe; sumsTs = ts; sumsS = s; }
            }
            if (mainP == null || sumsP == null)
                throw new ArgumentException("Шаблон не похож на заявление (не найдены абзацы дела/сумм).");

            // --- замены в абзаце дела ---
            string m2 = mainS;
            const string targetLawPreamble = "В соответствии с требованиями ст.ст. 30,33 Ф.З. № 229 от 02.10.2007 «Об исполнительном производстве», ";

            int idxProshu = m2.IndexOf("прошу принять", StringComparison.OrdinalIgnoreCase);
            if (idxProshu >= 0)
            {
                string afterProshu = m2.Substring(idxProshu);

                // 1. Номер дела/акта:
                string fullNum = p.FullNum;
                if (Regex.IsMatch(afterProshu, @"((?:судебный\s+приказ|судебного\s+приказа|исполнительный\s+документ|дело)?\s*)(?:ВС№|№|дело\s*№)?\s*[\w\d\-\/\.\s]+?(?=\s+от\s+\d{2}\.\d{2}\.\d{4})", RegexOptions.IgnoreCase))
                {
                    afterProshu = Regex.Replace(afterProshu, @"((?:судебный\s+приказ|судебного\s+приказа|исполнительный\s+документ|дело)?\s*)(?:ВС№|№|дело\s*№)?\s*[\w\d\-\/\.\s]+?(?=\s+от\s+\d{2}\.\d{2}\.\d{4})",
                        m => m.Groups[1].Value + fullNum, RegexOptions.IgnoreCase);
                }
                else
                {
                    afterProshu = Regex.Replace(afterProshu, @"(прошу принять\s+)(?:судебный приказ\s+)?.*?(?=\s+от\s+\d{2}\.\d{2}\.\d{4})",
                        m => m.Groups[1].Value + "судебный приказ " + fullNum, RegexOptions.IgnoreCase);
                }

                // 2. Дата дела (в afterProshu она единственная):
                afterProshu = Regex.Replace(afterProshu, @"(\s+от\s+)\d{2}\.\d{2}\.\d{4}",
                    m => m.Groups[1].Value + p.Date, RegexOptions.IgnoreCase);

                // 3. Должники
                afterProshu = Regex.Replace(afterProshu, @"(производство в отношении\s).*?(,\s*\(адрес объекта недвижимости\):\s)",
                    m => m.Groups[1].Value + debtorsClause + m.Groups[2].Value, RegexOptions.Singleline);

                // 4. Адрес
                afterProshu = Regex.Replace(afterProshu, @"(\(адрес объекта недвижимости\):\s).*?(,\s+а также осуществить)",
                    m => m.Groups[1].Value + addr + m.Groups[2].Value, RegexOptions.Singleline);

                m2 = targetLawPreamble + afterProshu;
            }

            if (!m2.Contains(debtorsClause)) throw new ArgumentException("Не удалось подставить должников (неожиданный шаблон).");
            if (!m2.Contains(addr)) throw new ArgumentException("Не удалось подставить адрес (неожиданный шаблон).");

            ApplyEdit(mainTs, 0, mainS.Length, m2);

            // жирный блок как в эталоне: от «прошу принять» до конца адреса
            int bStart, bLen;
            if (TryGetBoldSpan(m2, out bStart, out bLen))
                ApplyBoldRange(mainP, ns, mainTs, m2, bStart, bLen);

            ApplyEdit(sumsTs, 0, sumsS.Length, sumsText);

            WriteDoc(templatePath, outPath, doc);
        }

        // Пакетное исправление даты закона 229 (02.10.2007) во всех готовых заявлениях папки
        public int BatchFixLawDate(string folderPath, Action<string> log)
        {
            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException("Папка не найдена: " + folderPath);

            var files = new DirectoryInfo(folderPath).GetFiles("*.docx")
                .Where(x => !x.Name.StartsWith("~$")).ToList();

            if (files.Count == 0)
            {
                if (log != null) log("В папке нет файлов .docx для исправления.");
                return 0;
            }

            int fixedCount = 0;
            const string TargetLawHeader = "В соответствии с требованиями ст.ст. 30,33 Ф.З. № 229 от 02.10.2007 «Об исполнительном производстве»";

            foreach (var fi in files)
            {
                try
                {
                    XmlNamespaceManager ns;
                    XmlDocument doc = LoadDocXml(fi.FullName, out ns);
                    bool modified = false;

                    foreach (XmlNode pNode in doc.SelectNodes("//w:body//w:p", ns))
                    {
                        var pe = (XmlElement)pNode;
                        var ts = pe.SelectNodes(".//w:t", ns).Cast<XmlNode>().ToList();
                        if (ts.Count == 0) continue;
                        string s = ParaText(ts);

                        if ((s.Contains("229") || s.Contains("исполнительном производстве")) && s.Contains("прошу принять"))
                        {
                            int idxP = s.IndexOf("прошу принять", StringComparison.OrdinalIgnoreCase);
                            if (idxP >= 0)
                            {
                                string newS = TargetLawHeader + ", " + s.Substring(idxP);
                                if (newS != s)
                                {
                                    ApplyEdit(ts, 0, s.Length, newS);

                                    var tsUpdated = pe.SelectNodes(".//w:t", ns).Cast<XmlNode>().ToList();
                                    string sUpdated = ParaText(tsUpdated);
                                    int bStart, bLen;
                                    if (TryGetBoldSpan(sUpdated, out bStart, out bLen))
                                    {
                                        SetNodeText((XmlElement)tsUpdated[0], sUpdated);
                                        for (int i = 1; i < tsUpdated.Count; i++) SetNodeText((XmlElement)tsUpdated[i], "");
                                        ApplyBoldRange(pe, ns, tsUpdated, sUpdated, bStart, bLen);
                                    }

                                    modified = true;
                                }
                            }
                        }
                    }

                    if (modified)
                    {
                        string tmpPath = fi.FullName + ".tmp";
                        WriteDoc(fi.FullName, tmpPath, doc);
                        File.Delete(fi.FullName);
                        File.Move(tmpPath, fi.FullName);
                        fixedCount++;
                        if (log != null) log("✔ Исправлен: " + fi.Name);
                    }
                    else
                    {
                        if (log != null) log("• Без изменений (уже корректен): " + fi.Name);
                    }
                }
                catch (Exception ex)
                {
                    if (log != null) log("✖ Ошибка в файле " + fi.Name + ": " + ex.Message);
                }
            }

            if (log != null) log(string.Format("ИТОГО: обработано {0} файлов, исправлена дата закона в {1} файлах.", files.Count, fixedCount));
            return fixedCount;
        }

        public static GenParams ParseGenParams(string[] lines)
        {
            var p = new GenParams();
            foreach (string rawLine in lines)
            {
                if (string.IsNullOrWhiteSpace(rawLine)) continue;
                int eq = rawLine.IndexOf('=');
                if (eq <= 0) continue;
                string k = rawLine.Substring(0, eq).Trim().ToLowerInvariant();
                string v = rawLine.Substring(eq + 1).Trim();
                switch (k)
                {
                    case "prefix": p.NumPrefix = v; break;
                    case "num": p.Num = v; break;
                    case "date": p.Date = v; break;
                    case "addr": p.Addr = v; break;
                    case "rep": p.Rep = v; break;
                    case "solidary": p.Solidary = v == "1" || v == "да"; break;
                    case "osn": p.Osn = v; break;
                    case "gos": p.Gos = v; break;
                    case "sud": p.Sud = v; break;
                    case "osn_on": p.UseOsn = v != "0"; break;
                    case "gos_on": p.UseGos = v != "0"; break;
                    case "sud_on": p.UseSud = v != "0"; break;
                    default:
                        if (k.StartsWith("fio")) p.Debtors.Add(v);
                        break;
                }
            }
            return p;
        }
    }

    // ---------- точечная правка внутри абзаца ----------
    public struct Edit
    {
        public long Start;
        public long Len;
        public string Text;
        public string Log;
    }

    // ---------- интерактивный редактор сумм ----------
    public class SumSlot
    {
        public XmlElement Para;
        public List<XmlNode> Ts;
        public long Start;
        public long Len;
        public string Label;
        public string OldText;
        public string Prefill;
    }

    public class OpenedDoc
    {
        public string Path;
        public XmlDocument Doc;
        public XmlNamespaceManager Ns;
        public List<SumSlot> Slots = new List<SumSlot>();
    }

    public partial class Engine
    {
        internal const string WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // Диапазон жирного текста: от «прошу принять» до конца адреса,
        // НЕ включая «а также осуществить взыскание…» и реквизиты после него
        static bool TryGetBoldSpan(string s, out int start, out int len)
        {
            start = 0; len = 0;
            int st = s.IndexOf("прошу принять", StringComparison.Ordinal);
            if (st < 0) return false;
            int ea = s.IndexOf("а также осуществить", st, StringComparison.Ordinal);
            if (ea < 0) return false;
            int en = ea;
            while (en > st && (s[en - 1] == ' ' || s[en - 1] == ',')) en--;
            if (en <= st) return false;
            start = st;
            len = en - st;
            return true;
        }

        // Разбить текст абзаца на прогоны и сделать диапазон [start,start+len) жирным,
        // сохранив прочее форматирование первого прогона (шрифт, размер и т.д.)
        static void ApplyBoldRange(XmlElement p, XmlNamespaceManager ns, List<XmlNode> ts, string full, int start, int len)
        {
            if (ts.Count == 0) return;
            var r0 = ts[0].ParentNode as XmlElement;
            if (r0 == null || !r0.LocalName.Equals("r", StringComparison.Ordinal)) return;
            XmlDocument doc = p.OwnerDocument;

            var segs = new List<object[]>();
            if (start > 0) segs.Add(new object[] { full.Substring(0, start), false });
            segs.Add(new object[] { full.Substring(start, len), true });
            if (start + len < full.Length) segs.Add(new object[] { full.Substring(start + len), false });

            XmlElement rpr0 = r0.SelectSingleNode("./w:rPr", ns) as XmlElement;
            XmlNode anchor = r0;
            foreach (object[] sg in segs)
            {
                string txt = (string)sg[0];
                bool bold = (bool)sg[1];
                if (txt.Length == 0) continue;

                XmlElement nr = doc.CreateElement("w:r", WNS);
                if (rpr0 != null) nr.AppendChild(rpr0.Clone());
                XmlElement rprEl = nr.SelectSingleNode("./w:rPr", ns) as XmlElement;
                if (rprEl == null)
                {
                    rprEl = doc.CreateElement("w:rPr", WNS);
                    nr.PrependChild(rprEl);
                }
                XmlNode bel = rprEl.SelectSingleNode("./w:b", ns);
                XmlNode bcs = rprEl.SelectSingleNode("./w:bCs", ns);
                if (bold)
                {
                    // порядок в rPr по схеме: rFonts, b, bCs, ..., sz
                    XmlNode rfonts = rprEl.SelectSingleNode("./w:rFonts", ns);
                    if (bel == null)
                    {
                        var nb = doc.CreateElement("w:b", WNS);
                        if (rfonts != null) rprEl.InsertAfter(nb, rfonts); else rprEl.PrependChild(nb);
                        bel = nb;
                    }
                    if (bcs == null)
                    {
                        var nb2 = doc.CreateElement("w:bCs", WNS);
                        if (bel != null) rprEl.InsertAfter(nb2, bel); else rprEl.PrependChild(nb2);
                    }
                }
                else
                {
                    if (bel != null) bel.ParentNode.RemoveChild(bel);
                    if (bcs != null) bcs.ParentNode.RemoveChild(bcs);
                }
                XmlElement tEl = doc.CreateElement("w:t", WNS);
                tEl.SetAttribute("xml:space", "preserve");
                tEl.InnerText = txt;
                nr.AppendChild(tEl);

                p.InsertBefore(nr, anchor);
            }
            p.RemoveChild(r0);
        }

        public Action<string> Log = delegate { };

        public List<string> Changes = new List<string>();
        public List<string> Warnings = new List<string>();

        // ---------------- РЕГУЛЯРНЫЕ ВЫРАЖЕНИЯ ----------------
        static readonly Regex RxAddr = new Regex(
            @"(?s)\([^()]*юридический[^()]*\)[^;()]*?;\s*\([^()]*фактический[^()]*\)\s*");
        static readonly Regex RxKpp = new Regex(@"КПП\s*346101001");
        const string DoubleQuoteWord = "««ЛУКОЙЛ";

        static readonly Regex RxAmount = new Regex(
            @"в размере\s*(?<num>\d[\d\s\u00A0\u202F]*?)(?<dec>,\d{1,2})?\s*(?:руб\.?)?\s*\((?<words>[^()]*?)руб\.?\s*(?<kop>\d{1,2})\s*коп\.?\s*\)");
        static readonly Regex RxParen = new Regex(
            @"((?:в размере|руб\.)[\s\u00A0])(?<seg>[А-ЯЁ][^()]*?коп\.\))");
        static readonly Regex RxSpaceNum = new Regex(@"в размере(?=[(\d])");

        static readonly Regex RxNormal = new Regex(
            @"в размере\s*(?<num>\d[\d\s\u00A0\u202F]*)(?<dec>,\d{1,2})?\s*(?:руб\.?)?\s*\((?<inner>[^()]*)\)");
        static readonly Regex RxBroken = new Regex(
            @"в размере\s*(?<num>\d[\d\s\u00A0\u202F]*)(?<dec>,\d{1,2})?\s*руб\.?\s+(?<seg>[А-ЯЁ][^()]*?коп\.?\s*\))");

        // ---------------- ЧИСЛА <-> ПРОПИСЬ ----------------
        static readonly string[] OnesM = { "", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять" };
        static readonly string[] OnesF = { "", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять" };
        static readonly string[] Teens = { "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать" };
        static readonly string[] Tens = { "", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто" };
        static readonly string[] Hunds = { "", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот" };

        static string PluralForm(long g, string[] forms)
        {
            long m10 = g % 10, m100 = g % 100;
            if (m100 >= 11 && m100 <= 14) return forms[2];
            if (m10 == 1) return forms[0];
            if (m10 >= 2 && m10 <= 4) return forms[1];
            return forms[2];
        }

        static void AddGroup(List<string> parts, int g, bool fem)
        {
            int h = g / 100;
            int rest = g % 100;
            if (h > 0) parts.Add(Hunds[h]);
            if (rest >= 10 && rest <= 19) { parts.Add(Teens[rest - 10]); }
            else
            {
                int t = rest / 10, o = rest % 10;
                if (t >= 2) parts.Add(Tens[t]);
                if (o >= 1) parts.Add(fem ? OnesF[o] : OnesM[o]);
            }
        }

        public static string NumToRuWords(long n)
        {
            if (n < 0) return null;
            if (n == 0) return "ноль";
            var parts = new List<string>();
            long[] divs = { 1000000000L, 1000000L, 1000L };
            string[][] forms = {
                new[]{ "миллиард", "миллиарда", "миллиардов" },
                new[]{ "миллион", "миллиона", "миллионов" },
                new[]{ "тысяча", "тысячи", "тысяч" }
            };
            bool[] fems = { false, false, true };

            for (int i = 0; i < 3; i++)
            {
                long g = (n / divs[i]) % 1000;
                if (g == 0) continue;
                AddGroup(parts, (int)g, fems[i]);
                parts.Add(PluralForm(g, forms[i]));
            }
            int r = (int)(n % 1000);
            if (r > 0) AddGroup(parts, r, false);

            string res = string.Join(" ", parts.ToArray());
            res = char.ToUpper(res[0]) + res.Substring(1);
            return res;
        }

        static readonly Dictionary<string, long> RuUnits = new Dictionary<string, long>
        {
            {"один",1},{"одна",1},{"два",2},{"две",2},{"три",3},{"четыре",4},{"пять",5},
            {"шесть",6},{"семь",7},{"восемь",8},{"девять",9},{"десять",10},
            {"одиннадцать",11},{"двенадцать",12},{"тринадцать",13},{"четырнадцать",14},
            {"пятнадцать",15},{"шестнадцать",16},{"семнадцать",17},{"восемнадцать",18},
            {"девятнадцать",19},{"двадцать",20},{"тридцать",30},{"сорок",40},
            {"пятьдесят",50},{"шестьдесят",60},{"семьдесят",70},{"восемьдесят",80},
            {"девяносто",90},{"сто",100},{"двести",200},{"триста",300},{"четыреста",400},
            {"пятьсот",500},{"шестьсот",600},{"семьсот",700},{"восемьсот",800},{"девятьсот",900}
        };
        static readonly Dictionary<string, long> RuScales = new Dictionary<string, long>
        {
            {"тысяч",1000},{"тысяча",1000},{"тысячи",1000},
            {"миллион",1000000},{"миллиона",1000000},{"миллионов",1000000},
            {"миллиард",1000000000},{"миллиарда",1000000000},{"миллиардов",1000000000}
        };
        static readonly Dictionary<char, char> LatToCir = new Dictionary<char, char>
        {
            {'c','с'},{'e','е'},{'a','а'},{'x','х'},{'o','о'},{'p','р'},{'k','к'},
            {'m','м'},{'h','н'},{'y','у'},{'t','т'},{'b','в'}
        };

        public static bool TryWordsToNumber(string text, out long value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            var norm = text.ToLowerInvariant().Replace('ё', 'е');
            var sb = new StringBuilder();
            foreach (char ch0 in norm)
            {
                char ch = ch0;
                if (LatToCir.ContainsKey(ch)) ch = LatToCir[ch];
                sb.Append(ch);
            }
            string[] tokens = sb.ToString().Split(new[] { ' ', '\u00A0', '-', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            long total = 0, group = 0;
            bool any = false;
            foreach (string tok in tokens)
            {
                if (RuUnits.ContainsKey(tok)) { group += RuUnits[tok]; any = true; continue; }
                if (RuScales.ContainsKey(tok))
                {
                    if (!any) return false;
                    if (group == 0) group = 1;
                    total += group * RuScales[tok];
                    group = 0;
                    continue;
                }
                return false;
            }
            total += group;
            if (!any || total <= 0) return false;
            value = total;
            return true;
        }

        static bool TryParseDigits(string s, out long v)
        {
            string clean = Regex.Replace(s, @"[\s\u00A0\u202F]", "");
            return long.TryParse(clean, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out v);
        }

        // ---------------- РАБОТА С АБЗАЦЕМ ----------------
        static void SetNodeText(XmlElement t, string val)
        {
            t.InnerText = val;
            if (val != val.Trim()) t.SetAttribute("xml:space", "preserve");
        }

        static string ParaText(List<XmlNode> ts)
        {
            var sb = new StringBuilder();
            foreach (XmlNode n in ts) sb.Append(n.InnerText);
            return sb.ToString();
        }

        static void ApplyEdit(List<XmlNode> ts, long start, long len, string newText)
        {
            long pos = 0;
            bool doneFirst = false;
            foreach (XmlNode nodeObj in ts)
            {
                XmlElement node = (XmlElement)nodeObj;
                string txt = node.InnerText;
                long nodeStart = pos;
                long nodeEnd = pos + txt.Length;
                pos = nodeEnd;
                if (nodeEnd <= start || nodeStart >= start + len) continue;
                int localS = (int)Math.Max(0, start - nodeStart);
                int localE = (int)Math.Min(txt.Length, start + len - nodeStart);
                string pre = txt.Substring(0, localS);
                string post = localE < txt.Length ? txt.Substring(localE) : "";
                if (!doneFirst) { SetNodeText(node, pre + newText + post); doneFirst = true; }
                else SetNodeText(node, pre + post);
            }
        }

        static void ApplyEdits(List<XmlNode> ts, List<Edit> edits)
        {
            var sorted = edits.OrderBy(e => e.Start).ThenByDescending(e => e.Len).ToList();
            var final = new List<Edit>();
            long lastEnd = -1;
            foreach (var e in sorted)
            {
                if (e.Start >= lastEnd) { final.Add(e); lastEnd = e.Start + Math.Max(e.Len, 1); }
            }
            for (int i = final.Count - 1; i >= 0; i--)
                ApplyEdit(ts, final[i].Start, final[i].Len, final[i].Text);
        }

        // ---------------- РЕЖИМ 1: ИСПРАВЛЕНИЕ ----------------
        void ProcessParagraphFix(XmlElement p, XmlNamespaceManager ns)
        {
            bool edited = false;
            for (int round = 0; round < 6; round++)
            {
                var ts = p.SelectNodes(".//w:t", ns).Cast<XmlNode>().ToList();
                if (ts.Count == 0) return;
                string s = ParaText(ts);
                var edits = new List<Edit>();

                Match m = RxAddr.Match(s);
                if (m.Success)
                {
                    string head = m.Value.Trim();
                    if (head.Length > 60) head = head.Substring(0, 60);
                    edits.Add(new Edit { Start = m.Index, Len = m.Length, Text = "",
                        Log = "реквизиты: удалён блок \"" + head + "...\"" });
                }

                m = RxKpp.Match(s);
                if (m.Success)
                    edits.Add(new Edit { Start = m.Index, Len = m.Length, Text = "КПП 346001001",
                        Log = "КПП исправлен: 346101001 -> 346001001" });

                int qi = s.IndexOf(DoubleQuoteWord, StringComparison.Ordinal);
                if (qi >= 0)
                    edits.Add(new Edit { Start = qi, Len = 1, Text = "",
                        Log = "исправлена двойная кавычка ««ЛУКОЙЛ" });

                foreach (Match mm in RxAmount.Matches(s))
                {
                    Group gNum = mm.Groups["num"], gDec = mm.Groups["dec"],
                          gW = mm.Groups["words"], gKop = mm.Groups["kop"];
                    long digits, wordsVal;
                    bool okD = TryParseDigits(gNum.Value, out digits);
                    bool okW = TryWordsToNumber(gW.Value, out wordsVal);
                    if (!okD || !okW)
                    {
                        string frag = mm.Value; if (frag.Length > 70) frag = frag.Substring(0, 70);
                        Warnings.Add("не удалось сверить сумму \"" + frag + "\"");
                        continue;
                    }
                    if (digits != wordsVal)
                    {
                        Warnings.Add("РУБЛИ НЕ СОВПАДАЮТ: цифры=" + digits + ", прописью=" + wordsVal + " — оставлено без изменений");
                        continue;
                    }
                    string kopStr = int.Parse(gKop.Value).ToString("00");
                    long spanStart = gNum.Index;
                    long spanLen = gNum.Length + (gDec.Success ? gDec.Length : 0);
                    string origSpan = s.Substring((int)spanStart, (int)spanLen);
                    string numCore = gNum.Value.Trim(' ', '\u00A0', '\u202F');
                    string newSpan = null;
                    if (gDec.Success)
                    {
                        string curDec = gDec.Value.Substring(1);
                        if (curDec != kopStr) newSpan = numCore + "," + kopStr;
                    }
                    else if (kopStr != "00") newSpan = numCore + "," + kopStr;

                    if (newSpan != null)
                    {
                        Changes.Add("сумма: \"" + origSpan + "\" -> \"" + newSpan + "\" (копейки из прописи)");
                        edits.Add(new Edit { Start = spanStart, Len = spanLen, Text = newSpan, Log = "" });
                    }
                }

                foreach (Match mm in RxParen.Matches(s))
                {
                    Group gSeg = mm.Groups["seg"];
                    string head = gSeg.Value; if (head.Length > 40) head = head.Substring(0, 40);
                    Changes.Add("восстановлена скобка перед \"" + head + "...\"");
                    edits.Add(new Edit { Start = gSeg.Index, Len = 1, Text = "(" + gSeg.Value.Substring(0, 1), Log = "" });
                }

                foreach (Match mm in RxSpaceNum.Matches(s))
                {
                    int nextPos = mm.Index + mm.Length;
                    if (nextPos < s.Length)
                    {
                        Changes.Add("добавлен пробел: \"в размере...\"");
                        edits.Add(new Edit { Start = nextPos, Len = 1, Text = " " + s[nextPos], Log = "" });
                    }
                }

                if (s.Contains("Взыскатель просит") && !s.Contains("основного долга"))
                    Warnings.Add("в абзаце о взыскании НЕТ суммы основного долга (проверьте вручную)");

                if (edits.Count == 0) break;
                ApplyEdits(ts, edits);
                foreach (var e in edits) if (!string.IsNullOrEmpty(e.Log)) Changes.Add(e.Log);
                edited = true;
            }

            // восстановить жирный блок как в эталоне, если этот абзац переписывался
            if (edited)
            {
                var tsF = p.SelectNodes(".//w:t", ns).Cast<XmlNode>().ToList();
                string sf = ParaText(tsF);
                int bs, bl;
                if (TryGetBoldSpan(sf, out bs, out bl))
                {
                    // сначала собираем ВЕСЬ текст в первый прогон, иначе получится дубль
                    SetNodeText((XmlElement)tsF[0], sf);
                    for (int i = 1; i < tsF.Count; i++) SetNodeText((XmlElement)tsF[i], "");
                    ApplyBoldRange(p, ns, tsF, sf, bs, bl);
                }
            }
        }

        // ---------------- РЕЖИМ 2: ПРОПИСЬ ----------------
        static string BuildNewPropis(string numRaw, string decRaw, out long val)
        {
            val = 0;
            if (!TryParseDigits(numRaw, out val)) return null;
            if (val <= 0 || val > 999999999999L) return null;
            string words = NumToRuWords(val);
            if (words == null) return null;
            string kk = "00";
            if (!string.IsNullOrEmpty(decRaw)) kk = decRaw.Substring(1).PadRight(2, '0');
            return "(" + words + " руб. " + kk + " коп.)";
        }

        void ProcessParagraphPropis(XmlElement p, XmlNamespaceManager ns)
        {
            for (int round = 0; round < 6; round++)
            {
                var ts = p.SelectNodes(".//w:t", ns).Cast<XmlNode>().ToList();
                if (ts.Count == 0) return;
                string s = ParaText(ts);
                var edits = new List<Edit>();

                foreach (var rx in new[] { RxNormal, RxBroken })
                {
                    foreach (Match mm in rx.Matches(s))
                    {
                        Group gNum = mm.Groups["num"], gDec = mm.Groups["dec"];
                        long val;
                        string newText = BuildNewPropis(gNum.Value, gDec.Success ? gDec.Value : null, out val);
                        if (newText == null)
                        {
                            string frag = mm.Value; if (frag.Length > 70) frag = frag.Substring(0, 70);
                            Warnings.Add("не удалось записать прописью: \"" + frag + "\"");
                            continue;
                        }
                        long spanStart, spanLen;
                        if (rx == RxNormal)
                        {
                            Group gIn = mm.Groups["inner"];
                            spanStart = gIn.Index - 1;
                            spanLen = gIn.Length + 2;
                        }
                        else
                        {
                            Group gSeg = mm.Groups["seg"];
                            spanStart = gSeg.Index;
                            spanLen = gSeg.Length;
                        }
                        string oldTxt = s.Substring((int)spanStart, (int)spanLen);
                        if (oldTxt != newText)
                        {
                            Changes.Add(gNum.Value.Trim() + gDec.Value + " -> " + newText);
                            edits.Add(new Edit { Start = spanStart, Len = spanLen, Text = newText });
                        }
                    }
                }

                if (edits.Count == 0) break;
                ApplyEdits(ts, edits);
            }
        }

        // ---------------- РЕЖИМ 3: ИНТЕРАКТИВНЫЕ СУММЫ ----------------
        static string DetectLabel(string s, int beforeIndex, bool paraSolidary)
        {
            int best = -1;
            string lab = "Сумма";
            var kv = new[] {
                new { K = "основного долга", V = "Основной долг" },
                new { K = "госпошлины", V = "Госпошлина" },
                new { K = "судебных расходов", V = "Судебные расходы" }
            };
            foreach (var item in kv)
            {
                int ix = s.LastIndexOf(item.K, beforeIndex, StringComparison.Ordinal);
                if (ix >= 0 && ix > best) { best = ix; lab = item.V; }
            }
            if (paraSolidary) lab += " (солидарно)";
            return lab;
        }

        static string FormatGroup(string intPart)
        {
            if (intPart.Length <= 3) return intPart;
            var sb = new StringBuilder();
            for (int i = 0; i < intPart.Length; i++)
            {
                if (i > 0 && (intPart.Length - i) % 3 == 0) sb.Append(' ');
                sb.Append(intPart[i]);
            }
            return sb.ToString();
        }

        // разбор пользовательского ввода: "2500" / "2500,50" / "2500.5" / "2 500"
        public static bool TryParseAmount(string input, out string formatted, out long rubles, out string kk)
        {
            formatted = null; rubles = 0; kk = "00";
            if (input == null) return false;
            string compact = input.Trim().Replace(" ", "").Replace('\u00A0', ' ').Replace('.', ',');
            Match mDec = Regex.Match(compact, ",(\\d{1,2})$");
            Match mInt = Regex.Match(compact, "^(\\d{1,12})");
            if (!mInt.Success) return false;
            if (!long.TryParse(mInt.Groups[1].Value, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out rubles)) return false;
            if (rubles <= 0 || rubles > 999999999999L) return false;
            if (mDec.Success)
            {
                kk = mDec.Groups[1].Value.PadRight(2, '0');
                formatted = FormatGroup(mInt.Groups[1].Value) + "," + kk;
            }
            else
            {
                // целая сумма без копеек -> добавляем ,00
                formatted = FormatGroup(mInt.Groups[1].Value) + ",00";
            }
            return true;
        }

        public OpenedDoc OpenSums(string path)
        {
            string xmlText;
            using (FileStream fs = File.OpenRead(path))
            using (var z = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = z.GetEntry("word/document.xml");
                using (Stream es = entry.Open())
                using (var ms = new MemoryStream())
                {
                    es.CopyTo(ms); ms.Position = 0;
                    using (var reader = new StreamReader(ms, Encoding.UTF8))
                        xmlText = reader.ReadToEnd();
                }
            }
            var doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.LoadXml(xmlText);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            var od = new OpenedDoc { Path = path, Doc = doc, Ns = ns };
            foreach (XmlNode pNode in doc.SelectNodes("//w:body//w:p", ns))
            {
                XmlElement p = (XmlElement)pNode;
                var ts = p.SelectNodes(".//w:t", ns).Cast<XmlNode>().ToList();
                if (ts.Count == 0) continue;
                string s = ParaText(ts);
                bool solidary = s.Contains("солидарн");
                foreach (Match mm in RxNormal.Matches(s))
                {
                    Group gNum = mm.Groups["num"], gDec = mm.Groups["dec"], gIn = mm.Groups["inner"];
                    long spanStart = gNum.Index;
                    long spanLen = (gIn.Index - 1 + gIn.Length + 2) - spanStart;
                    var slot = new SumSlot
                    {
                        Para = p,
                        Ts = ts,
                        Start = spanStart,
                        Len = spanLen,
                        Label = DetectLabel(s, mm.Index, solidary),
                        OldText = s.Substring((int)spanStart, (int)spanLen),
                        Prefill = gNum.Value.Trim(' ', '\u00A0', '\u202F') + (gDec.Success ? gDec.Value : "")
                    };
                    od.Slots.Add(slot);
                }
            }
            return od;
        }

        public List<string> ApplyValues(OpenedDoc od, string[] inputs)
        {
            var msgs = new List<string>();
            var planned = new List<Tuple<SumSlot, string>>();
            for (int i = 0; i < od.Slots.Count; i++)
            {
                if (i >= inputs.Length) break;
                string inp = inputs[i];
                if (string.IsNullOrWhiteSpace(inp)) continue;
                string formatted; long rub; string kk;
                if (!TryParseAmount(inp, out formatted, out rub, out kk))
                {
                    msgs.Add("[!] НЕВЕРНОЕ ЧИСЛО: \"" + inp + "\" — пропущено (" + od.Slots[i].Label + ")");
                    continue;
                }
                string words = NumToRuWords(rub);
                string newText = formatted + " руб. (" + words + " руб. " + kk + " коп.)";
                SumSlot slot = od.Slots[i];
                if (slot.OldText == newText) continue;
                planned.Add(Tuple.Create(slot, newText));
                msgs.Add(slot.Label + ": " + newText);
            }
            // ВАЖНО: применяем правки внутри каждого абзаца справа налево,
            // чтобы координаты более левых правок оставались верными
            foreach (var grp in planned.GroupBy(t => t.Item1.Para))
            {
                List<XmlNode> ts = grp.First().Item1.Ts;
                foreach (var t in grp.OrderByDescending(x => x.Item1.Start))
                    ApplyEdit(ts, t.Item1.Start, t.Item1.Len, t.Item2);
            }
            foreach (var t in planned) t.Item1.OldText = t.Item2;
            return msgs;
        }

        public void WriteDoc(string srcPath, string dstPath, XmlDocument doc)
        {
            using (FileStream fsIn = File.OpenRead(srcPath))
            using (var zipIn = new ZipArchive(fsIn, ZipArchiveMode.Read))
            using (FileStream fsOut = File.Create(dstPath))
            using (var zipOut = new ZipArchive(fsOut, ZipArchiveMode.Create))
            {
                foreach (ZipArchiveEntry entry in zipIn.Entries)
                {
                    ZipArchiveEntry newEntry = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    if (entry.FullName == "word/document.xml")
                    {
                        var msOut = new MemoryStream();
                        var settings = new XmlWriterSettings
                        {
                            Encoding = new UTF8Encoding(false),
                            OmitXmlDeclaration = false
                        };
                        using (XmlWriter xw = XmlWriter.Create(msOut, settings))
                            doc.Save(xw);
                        byte[] bytes = msOut.ToArray();
                        Stream des = newEntry.Open();
                        des.Write(bytes, 0, bytes.Length);
                        des.Close();
                        msOut.Dispose();
                    }
                    else
                    {
                        using (Stream s1 = entry.Open())
                        using (Stream s2 = newEntry.Open())
                            s1.CopyTo(s2);
                    }
                }
            }
        }

        public void SaveOpenedDoc(OpenedDoc od)
        {
            string tmp = od.Path + ".tmp";
            WriteDoc(od.Path, tmp, od.Doc);
            try { File.Replace(tmp, od.Path, null); }
            catch { File.Copy(tmp, od.Path, true); File.Delete(tmp); }
        }

        // ---------------- ОБРАБОТКА ФАЙЛА ----------------
        void ProcessDoc(string srcPath, string dstPath, Action<XmlElement, XmlNamespaceManager> paraAction)
        {
            using (FileStream fsIn = File.OpenRead(srcPath))
            using (var zipIn = new ZipArchive(fsIn, ZipArchiveMode.Read))
            using (FileStream fsOut = File.Create(dstPath))
            using (var zipOut = new ZipArchive(fsOut, ZipArchiveMode.Create))
            {
                foreach (ZipArchiveEntry entry in zipIn.Entries)
                {
                    ZipArchiveEntry newEntry = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    if (entry.FullName == "word/document.xml")
                    {
                        using (Stream es = entry.Open())
                        using (var ms = new MemoryStream())
                        {
                            es.CopyTo(ms);
                            ms.Position = 0;
                            string xmlText;
                            using (var reader = new StreamReader(ms, Encoding.UTF8))
                                xmlText = reader.ReadToEnd();

                            var doc = new XmlDocument();
                            doc.PreserveWhitespace = true;
                            doc.LoadXml(xmlText);
                            var ns = new XmlNamespaceManager(doc.NameTable);
                            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                            foreach (XmlNode pNode in doc.SelectNodes("//w:body//w:p", ns))
                                paraAction((XmlElement)pNode, ns);

                            var msOut = new MemoryStream();
                            var settings = new XmlWriterSettings
                            {
                                Encoding = new UTF8Encoding(false),
                                OmitXmlDeclaration = false
                            };
                            using (XmlWriter xw = XmlWriter.Create(msOut, settings))
                                doc.Save(xw);

                            byte[] bytes = msOut.ToArray();
                            Stream des = newEntry.Open();
                            des.Write(bytes, 0, bytes.Length);
                            des.Close();
                            msOut.Dispose();
                        }
                    }
                    else
                    {
                        using (Stream s1 = entry.Open())
                        using (Stream s2 = newEntry.Open())
                            s1.CopyTo(s2);
                    }
                }
            }
        }

        void WriteReport(string path, StringBuilder sb)
        {
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        // ---------------- ЗАПУСК РЕЖИМА 1 ----------------
        public int RunFix(string projectDir, string outDirOverride = null)
        {
            string srcDir = Path.Combine(projectDir, "Исходные заявления");
            string outDir = outDirOverride ?? Path.Combine(projectDir, "Исправленные заявления");
            if (!Directory.Exists(srcDir))
            {
                Log("ОШИБКА: не найдена папка " + srcDir);
                return 1;
            }
            Directory.CreateDirectory(outDir);
            foreach (FileInfo old in new DirectoryInfo(outDir).GetFiles()) old.Delete();

            List<FileInfo> files = new DirectoryInfo(srcDir).GetFiles("*.docx").OrderBy(f => f.Name, StringComparer.Ordinal).ToList();
            var report = new StringBuilder();
            report.AppendLine("ОТЧЁТ ОБ ИСПРАВЛЕНИЯХ  (" + DateTime.Now.ToString("dd.MM.yyyy HH:mm") + ")");
            report.AppendLine("Источник: папка \"Исходные заявления\" (" + files.Count + " файлов). Результат: папка \"Исправленные заявления\"");
            report.AppendLine("Файлы без изменений тоже скопированы. Файлы не .docx пропускаются.");
            report.AppendLine(new string('=', 70));

            int done = 0, fixedCount = 0, warnCount = 0, errCount = 0;
            foreach (FileInfo f in files)
            {
                done++;
                Log(string.Format("[{0}/{1}] {2}", done, files.Count, f.Name));
                Changes.Clear(); Warnings.Clear();
                try
                {
                    ProcessDoc(f.FullName, Path.Combine(outDir, f.Name), ProcessParagraphFix);
                    if (Changes.Count == 0 && Warnings.Count == 0)
                    {
                        Log("    OK (без изменений)");
                        report.AppendLine("").AppendLine(f.Name + "  --  изменений не было");
                    }
                    else
                    {
                        if (Changes.Count > 0) { fixedCount++; Log("    ИСПРАВЛЕН (" + Changes.Count + ")"); }
                        else { warnCount++; Log("    ВНИМАНИЕ"); }
                        report.AppendLine("").AppendLine(f.Name);
                        foreach (string c in Changes.Distinct()) report.AppendLine("   [+] " + c);
                        foreach (string w in Warnings.Distinct()) report.AppendLine("   [!] " + w);
                    }
                }
                catch (Exception ex)
                {
                    errCount++;
                    Log("    ОШИБКА: " + ex.Message);
                    report.AppendLine("").AppendLine(f.Name + "  --  ОШИБКА: " + ex.Message);
                    string bad = Path.Combine(outDir, f.Name);
                    if (File.Exists(bad)) { try { File.Delete(bad); } catch { } }
                }
            }
            report.AppendLine("").AppendLine(new string('=', 70));
            WriteReport(Path.Combine(outDir, "ОТЧЁТ.txt"), report);
            Log("");
            Log(string.Format("Готово! Исправлено: {0}, предупреждения: {1}, ошибок: {2}, всего: {3}",
                fixedCount, warnCount, errCount, files.Count));
            Log("Результат: " + outDir);
            return 0;
        }

        // ---------------- ЗАПУСК РЕЖИМА 2 ----------------
        public int RunPropis(string target)
        {
            string workDir;
            List<FileInfo> files;
            bool singleFile = File.Exists(target);

            if (singleFile)
            {
                FileInfo fi = new FileInfo(target);
                workDir = fi.DirectoryName;
                files = new List<FileInfo> { fi };
            }
            else
            {
                if (!Directory.Exists(target))
                {
                    Log("ОШИБКА: не найдена папка " + target);
                    return 1;
                }
                workDir = target.TrimEnd('\\');
                files = new DirectoryInfo(workDir).GetFiles("*.docx").OrderBy(f => f.Name, StringComparer.Ordinal).ToList();
            }
            if (files.Count == 0) { Log("Нет .docx файлов для обработки."); return 1; }

            // резервные копии
            string backupDir = workDir + "_резерв_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            Directory.CreateDirectory(backupDir);
            foreach (FileInfo f in files) File.Copy(f.FullName, Path.Combine(backupDir, f.Name), true);
            Log("Резервные копии: " + backupDir);

            var report = new StringBuilder();
            report.AppendLine("ОТЧЁТ: ПРОПИСЬ СУММ ПО ЦИФРАМ  (" + DateTime.Now.ToString("dd.MM.yyyy HH:mm") + ")");
            report.AppendLine("Обработка: " + workDir + (singleFile ? " (один файл)" : ""));
            report.AppendLine(new string('=', 70));

            int done = 0, changed = 0, errCount = 0;
            foreach (FileInfo f in files)
            {
                done++;
                Log(string.Format("[{0}/{1}] {2}", done, files.Count, f.Name));
                Changes.Clear(); Warnings.Clear();
                string tmp = f.FullName + ".tmp";
                try
                {
                    ProcessDoc(f.FullName, tmp, ProcessParagraphPropis);
                    try { File.Replace(tmp, f.FullName, null); }
                    catch { File.Copy(tmp, f.FullName, true); File.Delete(tmp); }

                    if (Changes.Count > 0)
                    {
                        changed++;
                        Log("    ПРОПИСАНО (" + Changes.Count + ")");
                        report.AppendLine("").AppendLine(f.Name);
                        foreach (string c in Changes.Distinct()) report.AppendLine("   [+] " + c);
                        foreach (string w in Warnings.Distinct()) report.AppendLine("   [!] " + w);
                    }
                    else if (Warnings.Count > 0)
                    {
                        Log("    ВНИМАНИЕ");
                        report.AppendLine("").AppendLine(f.Name);
                        foreach (string w in Warnings.Distinct()) report.AppendLine("   [!] " + w);
                    }
                    else Log("    ок");
                }
                catch (Exception ex)
                {
                    errCount++;
                    Log("    ОШИБКА: " + ex.Message);
                    report.AppendLine("").AppendLine(f.Name + "  --  ОШИБКА: " + ex.Message);
                    if (File.Exists(tmp)) { try { File.Delete(tmp); } catch { } }
                }
            }

            string reportPath = Path.Combine(workDir, "ОТЧЁТ_ПРОПИСЬ.txt");
            WriteReport(reportPath, report);
            Log("");
            Log(string.Format("Готово! Изменённых файлов: {0} из {1}, ошибок: {2}", changed, files.Count, errCount));
            Log("Отчёт: " + reportPath);
            return 0;
        }
    }

    // ==================== ГЛАВНОЕ ОКНО ====================
    public class MainForm : Form
    {
        TextBox txtDir;
        Button btnBrowse, btnRun, btnNew, btnBatchFix;
        TextBox txtLog;
        ProgressBar progress;
        Label lblStatus;
        readonly Engine _engine = new Engine();

        public MainForm()
        {
            BuildUi();

            // при первом запуске создаём все нужные папки рядом с программой
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            if (!baseDir.EndsWith("\\")) baseDir += "\\";
            EnsureProjectFolders(baseDir);
            if (Directory.Exists(Path.Combine(baseDir, "Исходные заявления")))
                txtDir.Text = baseDir;

            AllowDrop = true;
            DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            };
            DragDrop += (s, e) =>
            {
                string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (paths != null && paths.Length > 0)
                {
                    string p = paths[0];
                    txtDir.Text = Directory.Exists(p) || File.Exists(p) ? p : txtDir.Text;
                }
            };
        }

        void BuildUi()
        {
            Text = "Помощник по заявлениям";
            Font = new Font("Segoe UI", 9.5f);
            StartPosition = FormStartPosition.CenterScreen;

            var wa = Screen.PrimaryScreen.WorkingArea;
            int fw = Math.Min(780, wa.Width - 16);
            int fh = Math.Min(660, wa.Height - 40);
            MinimumSize = new Size(fw, fh);
            Size = new Size(fw, fh);

            var lbl1 = new Label { Text = "Папка проекта:", Location = new Point(12, 12), AutoSize = true };
            txtDir = new TextBox
            {
                Location = new Point(12, 32),
                Width = 640,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            string here = AppDomain.CurrentDomain.BaseDirectory;
            if (Directory.Exists(Path.Combine(here, "Исходные заявления")) || Directory.Exists(Path.Combine(here, "Исправленные заявления")))
                txtDir.Text = here;
            btnBrowse = new Button
            {
                Text = "Обзор...",
                Location = new Point(658, 30),
                Size = new Size(84, 25),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnBrowse.Click += (s, e) =>
            {
                using (var fb = new FolderBrowserDialog())
                {
                    fb.Description = "Выберите папку проекта (в ней программа создаст папки «Исходные заявления», «Исправленные заявления», «Новые заявления»)";
                    if (Directory.Exists(txtDir.Text)) fb.SelectedPath = txtDir.Text;
                    if (fb.ShowDialog(this) == DialogResult.OK) txtDir.Text = fb.SelectedPath;
                }
            };

            btnRun = new Button
            {
                Text = "ОБРАБОТАТЬ ЗАЯВЛЕНИЯ  ▶\nшаг 1 — автоисправление документов   •   шаг 2 — указание сумм по каждому файлу",
                Location = new Point(12, 64),
                Size = new Size(730, 56),
                BackColor = Color.FromArgb(227, 242, 253),
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            btnRun.Click += (s, e) => RunFull();

            btnNew = new Button
            {
                Text = "+ СОЗДАТЬ НОВОЕ ЗАЯВЛЕНИЕ (генератор)",
                Location = new Point(12, 126),
                Size = new Size(340, 34),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnNew.Click += (s, e) => OpenGenerator();

            btnBatchFix = new Button
            {
                Text = "⚡ ИСПРАВИТЬ ДАТУ ЗАКОНА (02.10.2007) В ЗАЯВЛЕНИЯХ",
                Location = new Point(358, 126),
                Size = new Size(384, 34),
                BackColor = Color.FromArgb(255, 243, 224),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            btnBatchFix.Click += (s, e) => RunBatchFixLawDate();

            // --- описание папок ---
            var grpFolders = new GroupBox
            {
                Text = "Назначение папок",
                Location = new Point(12, 168),
                Size = new Size(730, 110),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Action<int, string, string> addFolderRow = (yy, name, desc) =>
            {
                grpFolders.Controls.Add(new Label { Text = name, Location = new Point(14, yy), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) });
                var dl = new Label { Text = desc, Location = new Point(190, yy + 1), AutoSize = true, ForeColor = Color.FromArgb(70, 70, 70), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
                grpFolders.Controls.Add(dl);
            };
            addFolderRow(20, "Исходные заявления", "положите сюда файлы Word, которые нужно исправить");
            addFolderRow(42, "Исправленные заявления", "здесь появятся готовые документы — останется вписать суммы");
            addFolderRow(64, "Новые заявления", "сюда сохраняются документы, созданные генератором");
            addFolderRow(86, "Шаблоны", "образец заявления для генератора (создаётся автоматически)");

            var btnOpenNewMain = new Button
            {
                Text = "📁 Открыть папку...",
                Location = new Point(585, 60),
                Size = new Size(130, 26),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnOpenNewMain.Click += (s, e) =>
            {
                string p = Path.Combine(txtDir.Text.Trim(), "Новые заявления");
                Directory.CreateDirectory(p);
                try { System.Diagnostics.Process.Start("explorer.exe", p); } catch { }
            };
            grpFolders.Controls.Add(btnOpenNewMain);

            Controls.Add(grpFolders);

            lblStatus = new Label
            {
                Text = "Готово к работе.",
                Location = new Point(12, 284),
                AutoSize = true,
                ForeColor = Color.DimGray,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            progress = new ProgressBar
            {
                Location = new Point(12, 306),
                Size = new Size(730, 18),
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Visible = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            txtLog = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font(FontFamily.GenericMonospace, 9f),
                Location = new Point(12, 332),
                Size = new Size(730, 224),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White
            };

            Controls.AddRange(new Control[] { lbl1, txtDir, btnBrowse, grpFolders, btnRun, btnNew, btnBatchFix, lblStatus, progress, txtLog });
        }

        // создаёт структуру папок проекта, если их ещё нет
        void EnsureProjectFolders(string dir)
        {
            try
            {
                foreach (string sub in new[] { "Исходные заявления", "Исправленные заявления", "Новые заявления", "Шаблоны" })
                    Directory.CreateDirectory(Path.Combine(dir, sub));

                string tplDir = Path.Combine(dir, "Шаблоны");

                // Если остался старый ошибочный шаблон — удаляем его
                string oldFake = Path.Combine(tplDir, "Шаблон заявления.docx");
                if (File.Exists(oldFake))
                {
                    try { File.Delete(oldFake); } catch { }
                }

                string masterTpl = Path.Combine(tplDir, "094729862.docx");
                if (!File.Exists(masterTpl))
                {
                    Engine.CreateDefaultTemplate(masterTpl);
                    AppendLog("Восстановлен эталонный образец заявления: " + masterTpl);
                }
            }
            catch { } // нет прав — пусть ошибка появится позже, при работе с папкой
        }

        void SetBusy(bool busy, string msg)
        {
            btnRun.Enabled = !busy;
            btnNew.Enabled = !busy;
            if (btnBatchFix != null) btnBatchFix.Enabled = !busy;
            btnBrowse.Enabled = !busy;
            progress.Visible = busy;
            lblStatus.Text = msg;
            lblStatus.ForeColor = busy ? Color.Firebrick : Color.DimGray;
        }

        async void RunBatchFixLawDate()
        {
            string dir = txtDir.Text;
            if (File.Exists(dir)) dir = Path.GetDirectoryName(dir);
            if (!Directory.Exists(dir))
            {
                MessageBox.Show("Укажите существующую папку проекта.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string targetDir = Path.Combine(dir, "Новые заявления");
            if (!Directory.Exists(targetDir)) targetDir = dir;

            var dr = MessageBox.Show(
                "Исправить дату закона № 229 (02.10.2007) и формулировку во всех заявлениях в папке:\n\n" +
                targetDir + "\n\n(Будут исправлены ошибочные даты закона, дата дела и номер останутся без изменений).\n\nНачать исправление?",
                "Пакетное исправление даты закона",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (dr != DialogResult.Yes) return;

            SetBusy(true, "Исправление даты закона в документах...");
            try
            {
                int count = 0;
                await Task.Run(() =>
                {
                    count = _engine.BatchFixLawDate(targetDir, AppendLog);
                });
                lblStatus.ForeColor = Color.Green;
                lblStatus.Text = string.Format("Готово. Исправлено файлов: {0}", count);
                MessageBox.Show(
                    string.Format("Пакетное исправление завершено!\n\nУспешно исправлено файлов: {0}\nПапка: {1}", count, targetDir),
                    "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Firebrick;
                lblStatus.Text = "Ошибка исправления.";
                AppendLog("ОШИБКА: " + ex.Message);
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false, "Готово к работе.");
            }
        }

        void AppendLog(string line)
        {
            if (InvokeRequired) { BeginInvoke(new Action(() => AppendLog(line))); return; }
            txtLog.AppendText(line + Environment.NewLine);
        }

        async void RunSafe(Action job)
        {
            if (!Directory.Exists(txtDir.Text) && !File.Exists(txtDir.Text))
            {
                MessageBox.Show("Укажите существующую папку проекта.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SetBusy(true, "Обработка...");
            try
            {
                await Task.Run(job);
                lblStatus.ForeColor = Color.Green;
                lblStatus.Text = "Готово. Смотрите журнал ниже.";
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Firebrick;
                lblStatus.Text = "Ошибка.";
                AppendLog("ОШИБКА: " + ex.Message);
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false, "Готово."); }
        }

        void OpenGenerator()
        {
            string dir = txtDir.Text;
            if (File.Exists(dir)) dir = Path.GetDirectoryName(dir);
            if (!Directory.Exists(dir))
            {
                MessageBox.Show("Укажите существующую папку проекта.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            EnsureProjectFolders(dir);
            using (var f = new GeneratorForm(dir, AppendLog))
            {
                f.ShowDialog(this);
            }
        }

        async void RunFull()
        {
            string dir = txtDir.Text;
            if (File.Exists(dir)) dir = Path.GetDirectoryName(dir);
            if (!Directory.Exists(dir))
            {
                MessageBox.Show("Укажите существующую папку.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            EnsureProjectFolders(dir);
            string srcDir = Path.Combine(dir, "Исходные заявления");
            if (Directory.GetFiles(srcDir, "*.docx").Length == 0)
            {
                MessageBox.Show(
                    "Папка \"Исходные заявления\" создана, но она пуста.\n\nПоложите туда документы .docx (заявления, которые нужно исправить) и запустите обработку снова.",
                    "Нет документов", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ---------- ШАГ 1: автоисправление ----------
            SetBusy(true, "Шаг 1 из 2: автоматическое исправление документов...");
            try
            {
                var engine = new Engine { Log = AppendLog };
                AppendLog("=== ШАГ 1: ИСПРАВЛЕНИЕ ДОКУМЕНТОВ ===");
                await Task.Run(() => engine.RunFix(dir));
            }
            catch (Exception ex)
            {
                AppendLog("ОШИБКА шага 1: " + ex.Message);
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetBusy(false, "Ошибка.");
                return;
            }

            // ---------- ШАГ 2: суммы по каждому документу ----------
            string target = Path.Combine(dir, "Исправленные заявления");
            lblStatus.Text = "Шаг 2 из 2: укажите суммы по документам...";
            AppendLog("");
            AppendLog("=== ШАГ 2: СУММЫ ПО ДОКУМЕНТАМ ===");
            using (var f = new SumEditorForm(target, AppendLog))
            {
                f.ShowDialog(this);
            }

            lblStatus.ForeColor = Color.Green;
            lblStatus.Text = "Готово! Результат: папка \"Исправленные заявления\".";
            SetBusy(false, "Готово. Результат в папке \"2_исправленные\".");
        }
    }

    // ==================== РЕДАКТОР СУММ ====================
    public class SumEditorForm : Form
    {
        readonly string _workDir;
        readonly Action<string> _log;
        readonly Engine _engine = new Engine();
        List<FileInfo> _files;
        int _idx = -1;
        OpenedDoc _od;
        TextBox[] _inputs;
        bool _backupMade;

        Label lblFile, lblPos, lblSaved, lblHint;
        Panel pnlScroll;
        TableLayoutPanel tlp;

        public SumEditorForm(string workDir, Action<string> log)
        {
            _workDir = workDir;
            _log = log;
            BuildUi();
            LoadList();
            if (_files.Count > 0) ShowFile(0);
        }

        void BuildUi()
        {
            Text = "Суммы по документам — " + _workDir;
            Font = new Font("Segoe UI", 10f);
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(720, 480);
            Size = new Size(780, 560);

            lblFile = new Label { Text = "", Location = new Point(12, 12), AutoSize = true, Font = new Font("Segoe UI", 11f, FontStyle.Bold), AutoEllipsis = true, Width = 500 };
            lblPos = new Label { Text = "", Location = new Point(660, 15), AutoSize = true, Anchor = AnchorStyles.Top | AnchorStyles.Right };

            lblHint = new Label
            {
                Text = "Введите суммы цифрами (можно с копейками через запятую). Пустое поле = оставить без изменений. Enter = записать и перейти дальше.",
                Location = new Point(12, 40),
                AutoSize = false,
                Size = new Size(740, 34),
                ForeColor = Color.DimGray
            };

            pnlScroll = new Panel
            {
                Location = new Point(12, 78),
                Size = new Size(740, 340),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            tlp = new TableLayoutPanel
            {
                ColumnCount = 2,
                AutoSize = true,
                Dock = DockStyle.Top,
                ColumnStyles = { new ColumnStyle(SizeType.Absolute, 330), new ColumnStyle(SizeType.Percent, 100) }
            };
            // ВАЖНО: добавляем таблицу с полями внутрь прокручиваемой панели
            pnlScroll.Controls.Add(tlp);

            lblSaved = new Label { Text = "", Location = new Point(14, 428), AutoSize = true, ForeColor = Color.Green, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };

            var btnPrev = new Button { Text = "◀ Предыдущий", Size = new Size(140, 34), Location = new Point(160, 455), Anchor = AnchorStyles.Bottom };
            var btnSkip = new Button { Text = "Пропустить ▶", Size = new Size(140, 34), Location = new Point(310, 455), Anchor = AnchorStyles.Bottom };
            var btnSaveNext = new Button { Text = "Записать и далее ▶", Size = new Size(200, 34), Location = new Point(460, 455), Anchor = AnchorStyles.Bottom, BackColor = Color.FromArgb(232, 245, 233) };

            btnPrev.Click += (s, e) => Step(-1);
            btnSkip.Click += (s, e) => Step(1);
            btnSaveNext.Click += (s, e) => SaveAndNext();
            AcceptButton = btnSaveNext;

            Controls.AddRange(new Control[] { lblFile, lblPos, lblHint, pnlScroll, lblSaved, btnPrev, btnSkip, btnSaveNext });
        }

        void LoadList()
        {
            _files = new DirectoryInfo(_workDir).GetFiles("*.docx").OrderBy(f => f.Name, StringComparer.Ordinal).ToList();
        }

        void RebuildSlots()
        {
            tlp.SuspendLayout();
            tlp.Controls.Clear();
            tlp.RowCount = 0;
            tlp.RowStyles.Clear();
            _inputs = new TextBox[_od.Slots.Count];
            for (int i = 0; i < _od.Slots.Count; i++)
            {
                SumSlot slot = _od.Slots[i];
                var lab = new Label
                {
                    Text = (i + 1) + ". " + slot.Label,
                    AutoSize = false,
                    Width = 320,
                    Height = 34,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = new Padding(3, 8, 3, 3),
                    Font = new Font("Segoe UI", 10.5f)
                };
                var tb = new TextBox
                {
                    Text = slot.Prefill,
                    Dock = DockStyle.Fill,
                    Margin = new Padding(3, 5, 25, 5),
                    Font = new Font("Segoe UI", 11f)
                };
                int row = i;
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
                tlp.RowCount++;
                _inputs[i] = tb;
                tlp.Controls.Add(lab, 0, row);
                tlp.Controls.Add(tb, 1, row);
            }
            if (_od.Slots.Count == 0)
            {
                tlp.RowCount = 1;
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.Controls.Add(new Label { Text = "В этом документе суммы не найдены.", AutoSize = true, ForeColor = Color.Firebrick }, 0, 0);
            }
            tlp.ResumeLayout();
            if (_inputs != null && _inputs.Length > 0) _inputs[0].Focus();
        }

        void ShowFile(int idx)
        {
            _idx = idx;
            _od = null;
            pnlScroll.Visible = false;
            try
            {
                FileInfo f = _files[idx];
                _od = _engine.OpenSums(f.FullName);
                lblFile.Text = f.Name;
                lblPos.Text = (idx + 1) + " / " + _files.Count;
                lblSaved.Text = "";
                RebuildSlots();
                pnlScroll.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть файл: " + ex.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void Step(int delta)
        {
            int ni = _idx + delta;
            if (ni < 0) ni = 0;
            if (ni >= _files.Count) ni = _files.Count - 1;
            if (ni != _idx) ShowFile(ni);
        }

        void EnsureBackup()
        {
            if (_backupMade) return;
            string bdir = _workDir.TrimEnd('\\') + "_резерв_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            Directory.CreateDirectory(bdir);
            foreach (FileInfo f in _files) File.Copy(f.FullName, Path.Combine(bdir, f.Name), true);
            _backupMade = true;
            _log("Резервные копии: " + bdir);
        }

        bool SaveCurrent()
        {
            if (_od == null || _inputs == null) return false;
            string[] vals = new string[_inputs.Length];
            for (int i = 0; i < _inputs.Length; i++) vals[i] = _inputs[i].Text;
            List<string> msgs = _engine.ApplyValues(_od, vals);
            foreach (string bad in msgs.Where(m => m.StartsWith("[!]")))
                MessageBox.Show(bad.Substring(3), "Проверьте ввод", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            List<string> good = msgs.Where(m => !m.StartsWith("[!]")).ToList();
            if (good.Count == 0) return false;
            try
            {
                EnsureBackup();
                _engine.SaveOpenedDoc(_od);
                foreach (string gmsg in good) _log(Path.GetFileName(_od.Path) + " → " + gmsg);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось сохранить: " + ex.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        void SaveAndNext()
        {
            bool saved = SaveCurrent();
            if (saved)
            {
                // перечитываем файл, чтобы поля показали записанное
                _engine.OpenSums(_files[_idx].FullName); // прогрев не нужен, просто перезагрузим ниже
                if (_idx < _files.Count - 1) ShowFile(_idx + 1);
                else { ShowFile(_idx); lblSaved.Text = "Сохранено ✔ Это был последний файл."; }
            }
            else
            {
                if (_idx < _files.Count - 1) Step(1);
            }
        }
    }

    // ==================== ГЕНЕРАТОР ЗАЯВЛЕНИЙ ====================
    public class GeneratorForm : Form
    {
        readonly string _projectDir;
        readonly Action<string> _log;
        readonly Engine _engine = new Engine();

        ComboBox cboNumPrefix;
        TextBox txtNum;
        DateTimePicker dtpDate;
        Label lblAddr, lblSums, lblRepHint;
        CheckBox chkSolid, chkOsn, chkGos, chkSud;
        TextBox txtOsn, txtGos, txtSud;
        RadioButton rbRepNone, rbRepFemale, rbRepMale, rbRepCustom;
        ComboBox cboRep;
        Label lblRepFio;
        TextBox txtRepFio, txtRepCustom;
        Label lblStreet, lblHouse, lblFlat;
        TextBox txtAddrStreet, txtAddrHouse, txtAddrFlat;
        CheckBox chkRoom;
        TextBox txtAddrRoom;
        Panel midPanel, bottomBar;
        Button btnAutofill;
        readonly Random _rnd = new Random();
        string _customTemplatePath;
        TextBox txtTemplate;

        static readonly string[] NumPrefixes = { "ВС№", "№", "дело №", "" };

        static readonly string[] SurnM = { "ИВАНОВ", "ПЕТРОВ", "СИДОРОВ", "КУЗНЕЦОВ", "СОКОЛОВ", "ПОПОВ", "ЛЕБЕДЕВ", "НОВИКОВ", "МОРОЗОВ", "ВОЛКОВ", "ЗАЙЦЕВ", "ЕФИМОВ" };
        static readonly string[] NameM = { "ИВАН", "АЛЕКСЕЙ", "СЕРГЕЙ", "ДМИТРИЙ", "НИКОЛАЙ", "ВЛАДИМИР", "ОЛЕГ", "ЮРИЙ", "АНДРЕЙ", "ПАВЕЛ" };
        static readonly string[] NameF = { "АННА", "ОЛЬГА", "МАРИЯ", "ЕЛЕНА", "НАТАЛЬЯ", "ИРИНА", "СВЕТЛАНА", "ЮЛИЯ", "КСЕНИЯ", "ТАТЬЯНА" };
        static readonly string[] PatrM = { "ИВАНОВИЧ", "ПЕТРОВИЧ", "АЛЕКСЕЕВИЧ", "СЕРГЕЕВИЧ", "НИКОЛАЕВИЧ", "ДМИТРИЕВИЧ", "ОЛЕГОВИЧ", "ВЛАДИМИРОВИЧ" };
        static readonly string[] PatrF = { "ИВАНОВНА", "ПЕТРОВНА", "АЛЕКСЕЕВНА", "СЕРГЕЕВНА", "НИКОЛАЕВНА", "ДМИТРИЕВНА", "ОЛЕГОВНА", "ВЛАДИМИРОВНА" };
        static readonly string[] Streets = { "ЛЕНИНА ул", "МИРА ул", "ЮБИЛЕЙНАЯ ул", "МОЛОДЕЖНАЯ ул", "ГАГАРИНА ул", "ПУШКИНА ул", "САДОВАЯ ул", "СТРОИТЕЛЕЙ ул", "ПОЛЕВАЯ ул", "5-Й КВ-Л ул" };
        NumericUpDown nudCount;
        TableLayoutPanel fioTlp;
        List<TextBox> fioInputs = new List<TextBox>();
        Label lblSaved;
        int _createdCount = 0;

        List<string> _loadedStreets = new List<string>();
        AutoCompleteStringCollection _streetAc = new AutoCompleteStringCollection();
        string _streetsFilePath;

        static readonly string[] DefaultStreets = {
            "11-Й КВ-Л ул", "3-Й МКР", "4-Й МКР", "5-Й КВ-Л ул", "5-Й МКР", "6-Й МКР", "7-Й МКР", "8-Й МКР",
            "БАЗАРОВА ул", "БОРОДИНСКАЯ ул", "БРАТСКАЯ ул", "ВАЛОВАЯ ул", "ВОИНОВ-ИНТЕРНАЦИОНАЛИСТОВ ул",
            "ВОЛГОГРАДСКАЯ ул", "ГАГАРИНА ул", "ГОРОХОВСКАЯ ул", "ДАЧНАЯ ул", "ДЕКАБРИСТОВ ул", "ДИМИТРОВА ул",
            "ДОС ул", "ЖЕЛЕЗНОДОРОЖНАЯ ул", "ЗЕЛЕНАЯ ул", "КАЛИНИНА ул", "КАХОВСКАЯ ул", "КИРОВА ул",
            "КОММУНАЛЬНАЯ ул", "КОМСОМОЛЬСКАЯ ул", "КОРОЛЕВА ул", "КОРОТКАЯ ул", "КРАСНОДОНСКАЯ ул",
            "КУЙБЫШЕВА ул", "ЛАЗАРЕВА ул", "ЛЕНИНА ул", "ЛЕРМОНТОВА ул", "ЛЕСНАЯ ул", "ЛОМОНОСОВА ул",
            "МАЯКОВСКОГО ул", "МЕТАЛЛУРГОВ ул", "МИРА ул", "МОЛОДЕЖНАЯ ул", "МУРМАНСКАЯ ул", "НАБЕРЕЖНАЯ ул",
            "НЕКРАСОВА ул", "НОВАЯ ул", "ОКТЯБРЬСКАЯ ул", "ОСТРОВСКОГО ул", "ПАРКОВАЯ ул", "ПЕРВОМАЙСКАЯ ул",
            "ПЕТРОВСКАЯ ул", "ПИОНЕРСКАЯ ул", "ПОБЕДЫ ул", "ПОДГОРНАЯ ул", "ПОЛЕВАЯ ул", "ПРИВОКЗАЛЬНАЯ ул",
            "ПРОЛЕТАРСКАЯ ул", "ПУШКИНА ул", "РАБОЧАЯ ул", "РАДИЩЕВА ул", "РЕСПУБЛИКАНСКАЯ ул",
            "РЯЗАНО-УРАЛЬСКАЯ ул", "САДОВАЯ ул", "САРАТОВСКАЯ ул", "СЕВЕРНАЯ ул", "СЕРОВА ул",
            "СОВЕТСКАЯ ул", "СПОРТИВНАЯ ул", "СТАХАНОВСКАЯ ул", "СТРОИТЕЛЕЙ ул", "ТЕКСТИЛЬНАЯ ул",
            "ТЕРЕШКОВОЙ ул", "ТИТОВА ул", "ТУРГЕНЕВА ул", "ФАБРИЧНАЯ ул", "ФЕДОСЕЕВА ул", "ФРУНЗЕ ул",
            "ЦИОЛКОВСКОГО ул", "ЧАПАЕВА ул", "ЧЕХОВА ул", "ЧКАЛОВА ул", "ШЕВЧЕНКО ул", "ЩОРСА ул",
            "ЮБИЛЕЙНАЯ ул"
        };

        void LoadStreets()
        {
            _streetsFilePath = Path.Combine(_projectDir, "улицы.txt");
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(_streetsFilePath))
            {
                try
                {
                    foreach (var line in File.ReadAllLines(_streetsFilePath, Encoding.UTF8))
                    {
                        string s = line.Trim();
                        if (s.Length > 0) set.Add(s);
                    }
                }
                catch { }
            }
            foreach (var s in DefaultStreets) set.Add(s);

            _loadedStreets = set.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
            try { File.WriteAllLines(_streetsFilePath, _loadedStreets, Encoding.UTF8); } catch { }

            _streetAc.Clear();
            _streetAc.AddRange(_loadedStreets.ToArray());
        }

        void RegisterStreet(string street)
        {
            if (string.IsNullOrWhiteSpace(street)) return;
            string s = street.Trim();
            if (!_loadedStreets.Any(x => x.Equals(s, StringComparison.OrdinalIgnoreCase)))
            {
                _loadedStreets.Add(s);
                _loadedStreets.Sort(StringComparer.OrdinalIgnoreCase);
                _streetAc.Add(s);
                try
                {
                    if (!string.IsNullOrEmpty(_streetsFilePath))
                        File.WriteAllLines(_streetsFilePath, _loadedStreets, Encoding.UTF8);
                }
                catch { }
            }
        }

        // Варианты представителей, разделенные по полу
        static readonly string[] FemalePhrases = {
            "мать: действующая за ребенка {ФИО}",
            "мать: действующая за детей {ФИО}",
            "действующая за своего несовершеннолетнего ребёнка {ФИО}",
            "действующая за себя и детей {ФИО}",
            "за неё действует законный представитель {ФИО}",
            "как законный представитель несовершеннолетнего {ФИО}"
        };

        static readonly string[] MalePhrases = {
            "отец: действующий за ребенка {ФИО}",
            "отец: действующий за детей {ФИО}",
            "действующий за своего несовершеннолетнего ребёнка {ФИО}",
            "действующий за себя и детей {ФИО}",
            "за него действует законный представитель {ФИО}",
            "как законный представитель несовершеннолетнего {ФИО}"
        };

        const string AddrPrefix = "403870, Волгоградская обл, Камышин г";

        const string EnChars = "`~qwertyuiop[]asdfghjkl;'zxcvbnm,./QWERTYUIOP{}ASDFGHJKL:\"ZXCVBNM<>?";
        const string RuChars = "ёЁйцукенгшщзхъфывапролджэячсмитьбю.ЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ,";

        public static void SwitchToRussian()
        {
            try
            {
                foreach (InputLanguage lang in InputLanguage.InstalledInputLanguages)
                {
                    if (lang.Culture.TwoLetterISOLanguageName.Equals("ru", StringComparison.OrdinalIgnoreCase))
                    {
                        InputLanguage.CurrentInputLanguage = lang;
                        break;
                    }
                }
            }
            catch { }
        }

        public static char EnToRuChar(char c)
        {
            int idx = EnChars.IndexOf(c);
            return idx >= 0 ? RuChars[idx] : c;
        }

        public static string FixLayout(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            var sb = new StringBuilder(input.Length);
            foreach (char c in input)
            {
                int idx = EnChars.IndexOf(c);
                sb.Append(idx >= 0 ? RuChars[idx] : c);
            }
            return sb.ToString();
        }

        public static string FixLatinHomoglyphs(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            const string lat = "ABCEHKMOPTXYabcehkmoptxy";
            const string cyr = "АВСЕНКМОРТХУАВСЕНКМОРТХУ";
            var sb = new StringBuilder(input.Length);
            foreach (char c in input)
            {
                int idx = lat.IndexOf(c);
                sb.Append(idx >= 0 ? cyr[idx] : c);
            }
            return sb.ToString();
        }

        public static void HookRussianLayout(TextBox tb, bool isAddressNumber = false)
        {
            if (tb == null) return;
            tb.Enter += (s, e) => SwitchToRussian();
            tb.KeyPress += (s, e) =>
            {
                char c = e.KeyChar;
                if (char.IsControl(c)) return;
                char converted = c;
                if (isAddressNumber)
                {
                    string hom = FixLatinHomoglyphs(c.ToString());
                    if (hom.Length > 0 && hom[0] != c) converted = hom[0];
                }
                if (converted == c)
                {
                    converted = EnToRuChar(c);
                }
                if (converted != c)
                {
                    e.KeyChar = converted;
                }
            };
            tb.TextChanged += (s, e) =>
            {
                string cur = tb.Text;
                if (string.IsNullOrEmpty(cur)) return;
                string fixedText = isAddressNumber ? FixLayout(FixLatinHomoglyphs(cur)) : FixLayout(cur);
                if (cur != fixedText)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = fixedText;
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
            };
        }

        public static string FormatStreet(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            string s = FixLayout(input.Trim().TrimEnd(',', '.')).ToUpperInvariant();
            var m = Regex.Match(s, @"^(?:УЛ\.|УЛ\.?\s+)(.+)$", RegexOptions.IgnoreCase);
            if (m.Success) s = m.Groups[1].Value.Trim();
            if (Regex.IsMatch(s, @"\s+УЛ\.?$", RegexOptions.IgnoreCase))
                return Regex.Replace(s, @"\s+УЛ\.?$", " ул", RegexOptions.IgnoreCase);
            return s + " ул";
        }

        public static string FormatHouse(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            string s = FixLayout(FixLatinHomoglyphs(input.Trim().TrimEnd(',', '.'))).ToUpperInvariant();
            var m = Regex.Match(s, @"^(?:Д\.\s*|Д\s+|ДОМ\s*(?:№|N)?\s*|№\s*)(.+)$", RegexOptions.IgnoreCase);
            if (m.Success) s = m.Groups[1].Value.Trim();
            return "дом №" + s;
        }

        public GeneratorForm(string projectDir, Action<string> log)
        {
            _projectDir = projectDir;
            _log = log;
            LoadStreets();
            BuildUi();
            RebuildFio();
        }

        Control AddRow(Control parent, int y, string caption, int labelW)
        {
            var l = new Label { Text = caption, Location = new Point(6, y + 3), AutoSize = false, Size = new Size(labelW - 10, 24), TextAlign = ContentAlignment.MiddleLeft };
            parent.Controls.Add(l);
            return l;
        }

        void BuildUi()
        {
            Text = "Генератор заявлений";
            Font = new Font("Segoe UI", 10f);
            StartPosition = FormStartPosition.CenterParent;
            KeyPreview = true;
            KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    Generate();
                }
            };

            var wa = Screen.PrimaryScreen.WorkingArea;
            int w = Math.Min(840, wa.Width - 16);
            int h = Math.Min(690, wa.Height - 40);
            MinimumSize = new Size(w, h);
            Size = new Size(w, h);

            // --- верхняя фикс-панель: номер дела, префикс, дата, число должников и выбор образца ---
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 100 };

            int y = 12;
            topPanel.Controls.Add(new Label { Text = "Номер дела / приказа:", Location = new Point(12, y + 4), AutoSize = true });

            cboNumPrefix = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(165, y),
                Width = 75,
                TabIndex = 0
            };
            cboNumPrefix.Items.AddRange(new object[] { "ВС№", "№", "дело №", "(пусто)" });
            cboNumPrefix.SelectedIndex = 0;
            topPanel.Controls.Add(cboNumPrefix);

            txtNum = new TextBox { Location = new Point(245, y), Width = 185, TabIndex = 1 };
            txtNum.Enter += (s, e) => SwitchToRussian();
            topPanel.Controls.Add(txtNum);

            topPanel.Controls.Add(new Label { Text = "Дата дела:", Location = new Point(445, y + 4), AutoSize = true });
            dtpDate = new DateTimePicker
            {
                Location = new Point(530, y),
                Width = 140,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "dd.MM.yyyy",
                MaxDate = new DateTime(2100, 12, 31),
                TabIndex = 2
            };
            topPanel.Controls.Add(dtpDate);

            y += 42;
            topPanel.Controls.Add(new Label { Text = "Должников:", Location = new Point(12, y + 3), AutoSize = true });
            nudCount = new NumericUpDown { Location = new Point(105, y), Minimum = 1, Maximum = 6, Value = 1, Width = 50, TabIndex = 3 };
            nudCount.ValueChanged += (s, e) => RebuildFio();
            topPanel.Controls.Add(nudCount);

            topPanel.Controls.Add(new Label { Text = "Образец (авто):", Location = new Point(175, y + 3), AutoSize = true, ForeColor = Color.DimGray });
            txtTemplate = new TextBox { Location = new Point(275, y), Width = 345, ReadOnly = true, TabStop = false };
            topPanel.Controls.Add(txtTemplate);
            var btnBrowseTemplate = new Button { Text = "Выбрать...", Location = new Point(626, y - 1), Size = new Size(100, 28), TabStop = false };
            btnBrowseTemplate.Click += (s, e) => ChooseTemplate();
            topPanel.Controls.Add(btnBrowseTemplate);

            // --- прокручиваемая середина ---
            midPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };

            fioTlp = new TableLayoutPanel
            {
                Location = new Point(12, 8),
                Width = 744,
                ColumnCount = 2,
                AutoSize = false,
                ColumnStyles = { new ColumnStyle(SizeType.Absolute, 140), new ColumnStyle(SizeType.Percent, 100) }
            };
            midPanel.Controls.Add(fioTlp);

            lblRepHint = new Label
            {
                Text = "Представитель / законный интерес (необязательно):",
                AutoSize = false,
                Size = new Size(744, 24),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            midPanel.Controls.Add(lblRepHint);

            rbRepNone = new RadioButton { Text = "Без представителя", Location = new Point(12, 0), AutoSize = true, Checked = true, TabIndex = 20 };
            rbRepFemale = new RadioButton { Text = "Женский (мать)", Location = new Point(175, 0), AutoSize = true, TabIndex = 21 };
            rbRepMale = new RadioButton { Text = "Мужской (отец)", Location = new Point(320, 0), AutoSize = true, TabIndex = 22 };
            rbRepCustom = new RadioButton { Text = "Свой вариант", Location = new Point(465, 0), AutoSize = true, TabIndex = 23 };

            rbRepNone.CheckedChanged += (s, e) => { if (rbRepNone.Checked) OnRepTypeChanged(); };
            rbRepFemale.CheckedChanged += (s, e) => { if (rbRepFemale.Checked) OnRepTypeChanged(); };
            rbRepMale.CheckedChanged += (s, e) => { if (rbRepMale.Checked) OnRepTypeChanged(); };
            rbRepCustom.CheckedChanged += (s, e) => { if (rbRepCustom.Checked) OnRepTypeChanged(); };

            midPanel.Controls.AddRange(new Control[] { rbRepNone, rbRepFemale, rbRepMale, rbRepCustom });

            cboRep = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 500,
                Visible = false,
                TabIndex = 24
            };
            midPanel.Controls.Add(cboRep);

            lblRepFio = new Label { Text = "ФИО ребёнка / представителя:", AutoSize = true, Visible = false };
            midPanel.Controls.Add(lblRepFio);
            txtRepFio = new TextBox { Width = 350, Visible = false, TabIndex = 25, CharacterCasing = CharacterCasing.Upper };
            HookRussianLayout(txtRepFio);
            midPanel.Controls.Add(txtRepFio);

            txtRepCustom = new TextBox { Width = 744, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, Visible = false, TabIndex = 26 };
            HookRussianLayout(txtRepCustom);
            midPanel.Controls.Add(txtRepCustom);

            lblAddr = new Label
            {
                Text = "Адрес объекта недвижимости (" + AddrPrefix + "):",
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            midPanel.Controls.Add(lblAddr);

            lblStreet = new Label { Text = "Улица:", AutoSize = true };
            midPanel.Controls.Add(lblStreet);
            txtAddrStreet = new TextBox { Width = 220, TabIndex = 30, CharacterCasing = CharacterCasing.Upper };
            txtAddrStreet.AutoCompleteCustomSource = _streetAc;
            txtAddrStreet.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtAddrStreet.AutoCompleteSource = AutoCompleteSource.CustomSource;
            HookRussianLayout(txtAddrStreet);
            txtAddrStreet.Leave += (s, e) =>
            {
                if (!string.IsNullOrWhiteSpace(txtAddrStreet.Text))
                    txtAddrStreet.Text = FormatStreet(txtAddrStreet.Text);
            };
            midPanel.Controls.Add(txtAddrStreet);

            lblHouse = new Label { Text = "Дом:", AutoSize = true };
            midPanel.Controls.Add(lblHouse);
            txtAddrHouse = new TextBox { Width = 70, TabIndex = 31, CharacterCasing = CharacterCasing.Upper };
            HookRussianLayout(txtAddrHouse, true);
            midPanel.Controls.Add(txtAddrHouse);

            lblFlat = new Label { Text = "Кв.:", AutoSize = true };
            midPanel.Controls.Add(lblFlat);
            txtAddrFlat = new TextBox { Width = 65, TabIndex = 32, CharacterCasing = CharacterCasing.Upper };
            HookRussianLayout(txtAddrFlat, true);
            midPanel.Controls.Add(txtAddrFlat);

            chkRoom = new CheckBox
            {
                Text = "комн.",
                AutoSize = true,
                TabIndex = 33,
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            chkRoom.CheckedChanged += (s, e) =>
            {
                txtAddrRoom.Enabled = chkRoom.Checked;
                if (chkRoom.Checked) txtAddrRoom.Focus();
                else txtAddrRoom.Text = "";
            };
            midPanel.Controls.Add(chkRoom);

            txtAddrRoom = new TextBox { Width = 65, Enabled = false, TabIndex = 34, CharacterCasing = CharacterCasing.Upper };
            HookRussianLayout(txtAddrRoom, true);
            midPanel.Controls.Add(txtAddrRoom);

            // Перенос солидарного взыскания вниз к суммам
            chkSolid = new CheckBox
            {
                Text = "Солидарное взыскание (взыскать в солидарном порядке со всех должников)",
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.DarkSlateBlue,
                TabIndex = 40
            };
            midPanel.Controls.Add(chkSolid);

            lblSums = new Label
            {
                Text = "Взыскать суммы:",
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            midPanel.Controls.Add(lblSums);

            chkOsn = new CheckBox { Text = "Основной долг", AutoSize = true, Checked = true, TabIndex = 41 };
            txtOsn = new TextBox { Width = 200, TabIndex = 42 };
            midPanel.Controls.Add(chkOsn); midPanel.Controls.Add(txtOsn);

            chkGos = new CheckBox { Text = "Госпошлина", AutoSize = true, Checked = false, TabIndex = 43 };
            txtGos = new TextBox { Width = 200, TabIndex = 44 };
            midPanel.Controls.Add(chkGos); midPanel.Controls.Add(txtGos);

            chkSud = new CheckBox { Text = "Судебные расходы", AutoSize = true, Checked = false, TabIndex = 45 };
            txtSud = new TextBox { Width = 200, TabIndex = 46 };
            midPanel.Controls.Add(chkSud); midPanel.Controls.Add(txtSud);

            // --- нижняя панель ---
            bottomBar = new Panel { Dock = DockStyle.Bottom, Height = 100, BackColor = Color.FromArgb(245, 246, 248) };
            bottomBar.Controls.Add(new Label
            {
                Text = "Сохранять в: " + Path.Combine(_projectDir, "Новые заявления"),
                Location = new Point(12, 8),
                AutoSize = true,
                ForeColor = Color.DimGray
            });
            lblSaved = new Label
            {
                Text = "",
                Location = new Point(14, 32),
                AutoSize = true,
                ForeColor = Color.Green
            };
            bottomBar.Controls.Add(lblSaved);

            btnAutofill = new Button
            {
                Text = "Автозаполнение (тест)",
                Size = new Size(170, 34),
                Location = new Point(12, 56),
                TabIndex = 51
            };
            btnAutofill.Click += (s, e) => Autofill();
            bottomBar.Controls.Add(btnAutofill);

            var btnOpenFolder = new Button
            {
                Text = "📁 Открыть папку с заявлениями",
                Size = new Size(240, 34),
                Location = new Point(190, 56),
                TabIndex = 52
            };
            btnOpenFolder.Click += (s, e) =>
            {
                string outDir = Path.Combine(_projectDir, "Новые заявления");
                Directory.CreateDirectory(outDir);
                try { System.Diagnostics.Process.Start("explorer.exe", outDir); }
                catch (Exception ex) { MessageBox.Show("Не удалось открыть папку: " + ex.Message); }
            };
            bottomBar.Controls.Add(btnOpenFolder);

            var btnGen = new Button
            {
                Text = "Создать заявление (Enter)",
                Size = new Size(220, 36),
                Location = new Point(bottomBar.Width - 234, 55),
                BackColor = Color.FromArgb(232, 245, 233),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                TabIndex = 50
            };
            btnGen.Click += (s, e) => Generate();
            bottomBar.Controls.Add(btnGen);

            Controls.Add(midPanel);
            Controls.Add(topPanel);
            Controls.Add(bottomBar);

            AcceptButton = btnGen;

            string initialTemplate = FindTemplate();
            txtTemplate.Text = initialTemplate ?? "(будет создан автоматически)";
        }

        void OnRepTypeChanged()
        {
            if (rbRepFemale.Checked)
            {
                cboRep.Items.Clear();
                cboRep.Items.AddRange(FemalePhrases);
                cboRep.SelectedIndex = 0;
            }
            else if (rbRepMale.Checked)
            {
                cboRep.Items.Clear();
                cboRep.Items.AddRange(MalePhrases);
                cboRep.SelectedIndex = 0;
            }
            RepositionMiddle();
        }

        void ChooseTemplate()
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Документы Word (*.docx)|*.docx|Все файлы (*.*)|*.*";
                ofd.Title = "Выберите файл-образец заявления (.docx)";
                string initDir = Path.Combine(_projectDir, "Шаблоны");
                if (!Directory.Exists(initDir)) initDir = _projectDir;
                ofd.InitialDirectory = initDir;
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    _customTemplatePath = ofd.FileName;
                    txtTemplate.Text = ofd.FileName;
                }
            }
        }

        void RebuildFio()
        {
            if (fioTlp == null || nudCount == null) return;
            fioTlp.SuspendLayout();
            fioTlp.Controls.Clear();
            fioTlp.RowStyles.Clear();
            fioTlp.RowCount = 0;
            fioInputs.Clear();
            int n = (int)nudCount.Value;
            for (int i = 0; i < n; i++)
            {
                int row = i;
                fioTlp.RowCount++;
                fioTlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
                string cap = n == 1 ? "Должник (ФИО):" : ("Должник " + (i + 1) + ":");
                fioTlp.Controls.Add(new Label { Text = cap, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3) }, 0, row);
                var tb = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 5, 25, 5), TabIndex = 10 + i, CharacterCasing = CharacterCasing.Upper };
                HookRussianLayout(tb);
                fioTlp.Controls.Add(tb, 1, row);
                fioInputs.Add(tb);
            }
            fioTlp.Height = n * 34 + 6;
            fioTlp.ResumeLayout();
            RepositionMiddle();
        }

        void RepositionMiddle()
        {
            int y = fioTlp.Bottom + 12;

            lblRepHint.Location = new Point(12, y);
            y += 26;

            rbRepNone.Location = new Point(12, y);
            rbRepFemale.Location = new Point(175, y);
            rbRepMale.Location = new Point(325, y);
            rbRepCustom.Location = new Point(475, y);
            y += 32;

            bool isGender = rbRepFemale.Checked || rbRepMale.Checked;
            bool isCustom = rbRepCustom.Checked;

            cboRep.Visible = isGender;
            lblRepFio.Visible = isGender;
            txtRepFio.Visible = isGender;
            txtRepCustom.Visible = isCustom;

            if (isGender)
            {
                cboRep.Location = new Point(12, y);
                y += 34;
                lblRepFio.Location = new Point(12, y + 4);
                txtRepFio.Location = new Point(255, y);
                y += 36;
            }
            else if (isCustom)
            {
                txtRepCustom.Location = new Point(12, y);
                y += 34;
            }

            lblAddr.Location = new Point(12, y);
            y += 26;

            lblStreet.Location = new Point(12, y + 3);
            txtAddrStreet.Location = new Point(68, y);

            lblHouse.Location = new Point(300, y + 3);
            txtAddrHouse.Location = new Point(345, y);

            lblFlat.Location = new Point(425, y + 3);
            txtAddrFlat.Location = new Point(458, y);

            chkRoom.Location = new Point(535, y + 3);
            txtAddrRoom.Location = new Point(598, y);
            y += 42;

            chkSolid.Location = new Point(12, y);
            y += 32;

            lblSums.Location = new Point(12, y);
            y += 26;

            chkOsn.Location = new Point(16, y); txtOsn.Location = new Point(190, y - 2); y += 34;
            chkGos.Location = new Point(16, y); txtGos.Location = new Point(190, y - 2); y += 34;
            chkSud.Location = new Point(16, y); txtSud.Location = new Point(190, y - 2);
        }

        string RndFio()
        {
            bool fem = _rnd.Next(2) == 0;
            int si = _rnd.Next(SurnM.Length);
            string surname = fem && !SurnM[si].EndsWith("А") ? SurnM[si] + "А" : SurnM[si];
            return surname + " " + (fem ? NameF[_rnd.Next(NameF.Length)] : NameM[_rnd.Next(NameM.Length)]) + " " +
                   (fem ? PatrF[_rnd.Next(PatrF.Length)] : PatrM[_rnd.Next(PatrM.Length)]);
        }

        string RndAmount(long min, long max, bool kopecks)
        {
            long v = _rnd.Next((int)min, (int)max + 1);
            if (_rnd.Next(2) == 1)
                return v.ToString("N0", System.Globalization.CultureInfo.InvariantCulture).Replace(',', ' ') + "," + _rnd.Next(0, 100).ToString("00");
            return v.ToString("N0", System.Globalization.CultureInfo.InvariantCulture).Replace(',', ' ');
        }

        void Autofill()
        {
            int n = _rnd.Next(3) == 0 ? _rnd.Next(2, 4) : 1;
            nudCount.Value = n;
            chkSolid.Checked = n > 1 && _rnd.Next(2) == 1;

            cboNumPrefix.SelectedIndex = _rnd.Next(cboNumPrefix.Items.Count);
            txtNum.Text = _rnd.Next(2) == 0
                ? _rnd.Next(100000000, 999999999).ToString()
                : string.Format("2-{0}-2024/2026", _rnd.Next(10, 99));
            dtpDate.Value = DateTime.Today.AddDays(-_rnd.Next(1, 400));

            fioInputs[0].Text = RndFio();
            for (int i = 1; i < fioInputs.Count; i++) fioInputs[i].Text = RndFio();

            int r = _rnd.Next(100);
            if (r < 40)
            {
                rbRepNone.Checked = true;
            }
            else if (r < 70)
            {
                rbRepFemale.Checked = true;
                cboRep.SelectedIndex = _rnd.Next(FemalePhrases.Length);
                txtRepFio.Text = RndFio();
            }
            else if (r < 90)
            {
                rbRepMale.Checked = true;
                cboRep.SelectedIndex = _rnd.Next(MalePhrases.Length);
                txtRepFio.Text = RndFio();
            }
            else
            {
                rbRepCustom.Checked = true;
                txtRepCustom.Text = "действующий в интересах несовершеннолетнего " + RndFio();
            }

            txtAddrStreet.Text = _loadedStreets.Count > 0 ? _loadedStreets[_rnd.Next(_loadedStreets.Count)] : "ЛЕНИНА ул";
            txtAddrHouse.Text = _rnd.Next(1, 60).ToString();
            txtAddrFlat.Text = _rnd.Next(1, 120).ToString();
            bool hasRoom = _rnd.Next(10) == 0;
            chkRoom.Checked = hasRoom;
            txtAddrRoom.Text = hasRoom ? _rnd.Next(1, 6).ToString() : "";

            int comboCase = _rnd.Next(100);
            bool useOsn, useGos = true, useSud;
            if (comboCase < 60) { useOsn = true; useSud = true; }
            else if (comboCase < 85) { useOsn = false; useSud = true; }
            else { useOsn = false; useSud = false; }

            chkOsn.Checked = useOsn;
            txtOsn.Text = useOsn ? RndAmount(1000, 85000, true) : "";
            chkGos.Checked = useGos;
            txtGos.Text = useGos ? RndAmount(400, 3000, false) : "";
            chkSud.Checked = useSud;
            txtSud.Text = useSud ? RndAmount(500, 5000, false) : "";
        }

        void Generate()
        {
            try
            {
                string template = FindTemplate();
                if (template == null)
                    throw new ArgumentException("Не найден образец заявления (.docx).\r\n\r\nПоложите файл-образец в папку \"Шаблоны\" или \"Исходные заявления\", либо нажмите «Выбрать...».");

                if (txtTemplate != null && (string.IsNullOrEmpty(txtTemplate.Text) || txtTemplate.Text.StartsWith("(")))
                    txtTemplate.Text = template;

                var p = new GenParams();
                int prefIdx = cboNumPrefix.SelectedIndex;
                p.NumPrefix = prefIdx >= 0 && prefIdx < NumPrefixes.Length ? NumPrefixes[prefIdx] : "";
                p.Num = txtNum.Text.Trim();
                p.Date = dtpDate.Value.ToString("dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);

                // фраза о представителе
                if (rbRepFemale.Checked || rbRepMale.Checked)
                {
                    string fio = FixLayout(txtRepFio.Text.Trim());
                    if (fio.Length == 0)
                        throw new ArgumentException("Выбран вариант представительства — впишите ФИО ребёнка/представителя.");
                    string templatePhrase = cboRep.SelectedItem != null ? cboRep.SelectedItem.ToString() : "";
                    if (templatePhrase.StartsWith("мать: ")) templatePhrase = templatePhrase.Substring(6);
                    if (templatePhrase.StartsWith("отец: ")) templatePhrase = templatePhrase.Substring(6);
                    p.Rep = templatePhrase.Replace("{ФИО}", fio.ToUpperInvariant());
                }
                else if (rbRepCustom.Checked)
                {
                    p.Rep = FixLayout(txtRepCustom.Text.Trim());
                    if (p.Rep.Length == 0)
                        throw new ArgumentException("Выбран «свой вариант» — впишите фразу полностью.");
                }
                else
                {
                    p.Rep = "";
                }

                // сборка адреса
                string street = FormatStreet(txtAddrStreet.Text);
                RegisterStreet(street);
                string house = txtAddrHouse.Text.Trim();
                string flat = FixLatinHomoglyphs(FixLayout(txtAddrFlat.Text.Trim())).ToUpperInvariant();
                string room = chkRoom.Checked ? FixLatinHomoglyphs(FixLayout(txtAddrRoom.Text.Trim())).ToUpperInvariant() : "";

                if (string.IsNullOrEmpty(street) && string.IsNullOrEmpty(house))
                    throw new ArgumentException("Впишите адрес (улицу и дом).");

                var addrParts = new List<string>();
                addrParts.Add(AddrPrefix);
                if (!string.IsNullOrEmpty(street)) addrParts.Add(street);
                if (!string.IsNullOrEmpty(house))
                {
                    string h = FormatHouse(house);
                    addrParts.Add(h);
                }
                if (!string.IsNullOrEmpty(flat))
                {
                    string cleanF = Regex.Replace(flat, @"^кв\.?\s*", "", RegexOptions.IgnoreCase).Trim().ToUpperInvariant();
                    addrParts.Add("кв. " + cleanF);
                }
                if (chkRoom.Checked && !string.IsNullOrEmpty(room))
                {
                    string cleanR = Regex.Replace(room, @"^комн\.?\s*", "", RegexOptions.IgnoreCase).Trim().ToUpperInvariant();
                    addrParts.Add("комн. " + cleanR);
                }
                p.Addr = string.Join(", ", addrParts);

                p.Solidary = chkSolid.Checked;
                foreach (var tb in fioInputs) p.Debtors.Add(FixLayout(tb.Text.Trim()).ToUpperInvariant());
                // Проверка и подготовка сумм
                bool useOsn = chkOsn.Checked && !string.IsNullOrWhiteSpace(txtOsn.Text);
                bool useGos = chkGos.Checked && !string.IsNullOrWhiteSpace(txtGos.Text);
                bool useSud = chkSud.Checked && !string.IsNullOrWhiteSpace(txtSud.Text);

                if (chkOsn.Checked && string.IsNullOrWhiteSpace(txtOsn.Text))
                    throw new ArgumentException("Вы отметили «Основной долг», но поле суммы пустое.\n\nВведите сумму (например: 15000,00) или снимите галочку.");
                if (chkGos.Checked && string.IsNullOrWhiteSpace(txtGos.Text))
                    throw new ArgumentException("Вы отметили «Госпошлина», но поле суммы пустое.\n\nВведите сумму (например: 1259,24) или снимите галочку.");
                if (chkSud.Checked && string.IsNullOrWhiteSpace(txtSud.Text))
                    throw new ArgumentException("Вы отметили «Судебные расходы», но поле суммы пустое.\n\nВведите сумму (например: 1500,00) или снимите галочку.");

                if (!useOsn && !useGos && !useSud)
                    throw new ArgumentException("Укажите хотя бы одну сумму для взыскания (например, сумму основного долга).");

                p.UseOsn = useOsn; p.Osn = txtOsn.Text.Trim();
                p.UseGos = useGos; p.Gos = txtGos.Text.Trim();
                p.UseSud = useSud; p.Sud = txtSud.Text.Trim();

                string outDir = Path.Combine(_projectDir, "Новые заявления");
                Directory.CreateDirectory(outDir);
                string safeFileName = GenParams.SanitizeFileName(p.Num);
                string outPath = Path.Combine(outDir, safeFileName + ".docx");

                _engine.GenerateApplication(template, outPath, p);

                _createdCount++;
                lblSaved.Text = string.Format("✔ Сохранено: {0}.docx  (Всего в этой сессии: {1})", safeFileName, _createdCount);
                _log("Сгенерировано заявление: " + outPath);

                txtNum.Text = "";
                foreach (var tb in fioInputs) tb.Text = "";
                txtRepFio.Text = "";
                txtAddrStreet.Text = "";
                txtAddrHouse.Text = "";
                txtAddrFlat.Text = "";
                chkRoom.Checked = false;
                txtAddrRoom.Text = "";
                txtNum.Focus();
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(ex.Message, "Проверьте данные", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        string FindTemplate()
        {
            // 0. Если пользователь явно выбрал файл в интерфейсе
            if (!string.IsNullOrEmpty(_customTemplatePath) && File.Exists(_customTemplatePath))
                return _customTemplatePath;

            // 1. Папка "Шаблоны" или "Шаблон" в проекте
            foreach (string sub in new[] { "Шаблоны", "Шаблон" })
            {
                string d = Path.Combine(_projectDir, sub);
                if (!Directory.Exists(d)) continue;

                // Удаляем старый ошибочный "Шаблон заявления.docx", если он остался от прежней версии
                string oldFake = Path.Combine(d, "Шаблон заявления.docx");
                if (File.Exists(oldFake))
                {
                    try { File.Delete(oldFake); } catch { }
                }

                var files = new DirectoryInfo(d).GetFiles("*.docx")
                    .Where(x => !x.Name.StartsWith("~$"))
                    .OrderByDescending(x => x.Name.StartsWith("094729862", StringComparison.OrdinalIgnoreCase))
                    .ThenBy(x => x.Name, StringComparer.Ordinal).ToList();
                if (files.Count > 0) return files[0].FullName;
            }

            // 2. Папки обработанных и исходных документов
            foreach (string sub in new[] { "Исправленные заявления", "Исходные заявления" })
            {
                string d = Path.Combine(_projectDir, sub);
                if (!Directory.Exists(d)) continue;
                var f = new DirectoryInfo(d).GetFiles("*.docx")
                    .Where(x => !x.Name.StartsWith("~$"))
                    .OrderBy(x => x.Name, StringComparer.Ordinal).FirstOrDefault();
                if (f != null) return f.FullName;
            }

            // 3. Документы в корне папки проекта (предпочтительно со словом шаблон/образец)
            if (Directory.Exists(_projectDir))
            {
                var f = new DirectoryInfo(_projectDir).GetFiles("*.docx")
                    .Where(x => !x.Name.StartsWith("~$"))
                    .OrderByDescending(x => x.Name.StartsWith("094729862", StringComparison.OrdinalIgnoreCase))
                    .ThenByDescending(x => x.Name.IndexOf("шаблон", StringComparison.OrdinalIgnoreCase) >= 0 || x.Name.IndexOf("образец", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ThenBy(x => x.Name, StringComparer.Ordinal).FirstOrDefault();
                if (f != null) return f.FullName;
            }

            // 4. Папка "Шаблоны" рядом с exe программы
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            if (!string.IsNullOrEmpty(exeDir) && Directory.Exists(exeDir) && !exeDir.TrimEnd('\\', '/').Equals(_projectDir.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase))
            {
                string d = Path.Combine(exeDir, "Шаблоны");
                if (Directory.Exists(d))
                {
                    var files = new DirectoryInfo(d).GetFiles("*.docx")
                        .Where(x => !x.Name.StartsWith("~$"))
                        .OrderByDescending(x => x.Name.StartsWith("094729862", StringComparison.OrdinalIgnoreCase))
                        .ThenBy(x => x.Name, StringComparer.Ordinal).ToList();
                    if (files.Count > 0) return files[0].FullName;
                }
            }

            // 5. Если ничего не найдено на новом ПК — автоматически восстанавливаем эталонный образец 094729862.docx из встроенных ресурсов!
            try
            {
                string tplDir = Path.Combine(_projectDir, "Шаблоны");
                Directory.CreateDirectory(tplDir);
                string autoPath = Path.Combine(tplDir, "094729862.docx");
                Engine.CreateDefaultTemplate(autoPath);
                _log("Восстановлен эталонный образец заявления: " + autoPath);
                if (File.Exists(autoPath)) return autoPath;
            }
            catch (Exception ex)
            {
                _log("Предупреждение: не удалось восстановить эталонный шаблон: " + ex.Message);
            }

            return null;
        }
    }

    public static class Program
    {
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);

        [STAThread]
        public static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // командная строка: app.exe fix <проект>   |   app.exe propis <папка|файл>
            if (args.Length >= 1)
            {
                AttachConsole(-1);
                string mode = args[0].ToLowerInvariant();
                string dir = args.Length >= 2 ? args[1] : Directory.GetCurrentDirectory();
                var engine = new Engine { Log = l => Console.WriteLine(l) };
                try
                {
                    if (mode == "fix") return engine.RunFix(dir);
                    if (mode == "propis") return engine.RunPropis(dir);
                    if (mode == "fixdate")
                    {
                        int fixedCount = engine.BatchFixLawDate(dir, l => Console.WriteLine(l));
                        return fixedCount >= 0 ? 0 : 1;
                    }
                    if (mode == "scan")
                    {
                        OpenedDoc od = engine.OpenSums(dir);
                        Console.WriteLine("Файл: " + od.Path);
                        for (int i = 0; i < od.Slots.Count; i++)
                            Console.WriteLine("[" + (i + 1) + "] " + od.Slots[i].Label + " | текущее: " + od.Slots[i].Prefill + " | (" + od.Slots[i].OldText + ")");
                        return 0;
                    }
                    if (mode == "uitest")
                    {
                        var form = new SumEditorForm(args[1], l => Console.WriteLine(l));
                        form.Show();
                        Application.DoEvents();
                        var fi = typeof(SumEditorForm).GetField("_inputs", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        var inputs = fi != null ? fi.GetValue(form) as TextBox[] : null;
                        Console.WriteLine("полей ввода: " + (inputs == null ? -1 : inputs.Length));
                        var fp = typeof(SumEditorForm).GetField("pnlScroll", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        var panel = fp != null ? fp.GetValue(form) as Panel : null;
                        if (panel != null && panel.Controls.Count > 0)
                        {
                            Console.WriteLine("панель: контролов=" + panel.Controls.Count + ", в таблице=" + panel.Controls[0].Controls.Count);
                            foreach (Control c in panel.Controls[0].Controls)
                                Console.WriteLine("  - " + c.GetType().Name + ": " + c.Text);
                        }
                        else Console.WriteLine("ПАНЕЛЬ ПУСТА!");
                        form.Close();
                        form.Dispose();
                        return 0;
                    }
                    if (mode == "gen" && args.Length >= 4)
                    {
                        // gen <шаблон> <выход> <файл_параметров>
                        GenParams gp;
                        string[] lines = File.ReadAllLines(args[3], Encoding.UTF8);
                        gp = Engine.ParseGenParams(lines);
                        engine.GenerateApplication(args[1], args[2], gp);
                        Console.WriteLine("СОЗДАНО: " + args[2]);
                        return 0;
                    }
                    if (mode == "setsums" && args.Length >= 3)
                    {
                        OpenedDoc od = engine.OpenSums(args[1]);
                        string[] vals = args[2].Split(';');
                        List<string> msgs = engine.ApplyValues(od, vals);
                        foreach (string m in msgs) Console.WriteLine(m);
                        if (msgs.Any(x => !x.StartsWith("[!]")))
                        {
                            engine.SaveOpenedDoc(od);
                            Console.WriteLine("СОХРАНЕНО: " + od.Path);
                        }
                        return 0;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ОШИБКА: " + ex.Message);
                    return 2;
                }
                Console.WriteLine("Использование: app fix <папка> | app propis <папка>");
                return 2;
            }

            Application.Run(new MainForm());
            return 0;
        }
    }
}
