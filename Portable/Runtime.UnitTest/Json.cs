using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeExtension;

namespace Runtime.UnitTest
{
    [TestClass]
    public class Json
    {
        [TestMethod]
        public void TestMethod1()
        {
            RequestMessageBody body = new RequestMessageBody();
            body.Actions = new List<ActionInfo>();
            body.Actions.Add(new ActionInfo() { Id = 1, ActionType = ActionType.Method, Name = "Trace", ArgumentInfo = new ArgumentInfo() { Arguments = new object[] { "Hello" } } });
            body.Actions.Add(new ActionInfo() { Id = 2, ActionType = ActionType.Method, Name = "Trace", ArgumentInfo = new ArgumentInfo() { Arguments = new object[] { "Hello" } } });

            string str = Utility.ToJsonString(body);
            Assert.IsTrue(true, str);
        }
    }
}
