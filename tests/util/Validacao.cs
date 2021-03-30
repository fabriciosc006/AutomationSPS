using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;

namespace SiggaPS.tests.util
{
    class Validacao
    {
        public bool ValidaElemVisivel(IWebElement elemento)
        {
            try
            {
                Util util = new Util();
                if (elemento.Displayed)
                {
                    util.HighlightElementPassou(elemento);
                    return true;
                }
                else
                {
                    util.HighlightElementFalhou(elemento);
                    return false;
                }

            }
            catch (Exception)
            {
                return false;
            }
        }
        public bool ValidaElemEnable(IWebElement elemento)
        {
            try
            {
                Util util = new Util();
                if (elemento.Enabled)
                {
                    util.HighlightElementPassou(elemento);
                    return true;
                }
                else
                {
                    util.HighlightElementFalhou(elemento);
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
        public bool ValidaElemContainsText(IWebElement elemento, string resultadoEsperado)
        {
            try
            {
                Util util = new Util();
                string mes = elemento.Text;
                if (elemento.Text.Contains(resultadoEsperado) && resultadoEsperado != "")
                {
                    util.HighlightElementPassou(elemento);
                    return true;
                }
                else
                {
                    util.HighlightElementFalhou(elemento);
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ValidaEqualsText(string resultadoEsperado1, string resultadoEsperado2)
        {
            try
            {
                return resultadoEsperado1.Equals(resultadoEsperado2);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ValidaElemEqualsText(IWebElement elemento, string resultadoEsperado)
        {
            try
            {
                Util util = new Util();
                if (elemento.Text.Equals(resultadoEsperado))
                {
                    util.HighlightElementPassou(elemento);
                    return true;
                }
                else if (elemento.Text != resultadoEsperado)
                {
                    util.HighlightElementFalhou(elemento);
                    return false;
                }
                else
                {
                    util.HighlightElementFalhou(elemento);
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ValidaAttributeValueText(IWebElement elemento, string resultadoEsperado)
        {
            try
            {
                Util util = new Util();
                if (elemento.GetAttribute("innerText").Contains(resultadoEsperado) && resultadoEsperado != "")
                {
                    util.HighlightElementPassou(elemento);
                    return true;
                }
                else
                {
                    util.HighlightElementFalhou(elemento);
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ValidaContainsText(string resultadoEsperado1, string resultadoEsperado2)
        {
            try
            {
                return resultadoEsperado1.Contains(resultadoEsperado2);
            }
            catch (Exception)
            {
                return false;
            }
        }

    }
}

