using System;
using OpenQA.Selenium;

namespace SOSD_AutomateSR
{
    internal class SelectElement
    {
        private IWebElement selectElement;

        public SelectElement(IWebElement selectElement)
        {
            this.selectElement = selectElement;
        }

        internal void SelectByValue(string v)
        {
            throw new NotImplementedException();
        }
    }
}