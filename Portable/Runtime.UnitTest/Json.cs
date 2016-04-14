using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeExtension;
using Newtonsoft.Json.Linq;

namespace Runtime.UnitTest
{
    [TestClass]
    public class Json
    {
        public TestContext TestContext
        {
            get;
            set;
        }

		[TestMethod]
		public void TestJToken()
		{
			JToken token = JToken.FromObject("abc");
			object obj = token.ToObject<object>();
			TestContext.WriteLine("obj={0}", obj);
			token = JToken.FromObject(123);
			obj = token.ToObject<object>();
			TestContext.WriteLine("obj={0}", obj);
		}

        [TestMethod]
        public void SerializeRequestMessageBody()
        {
            RequestMessageBody body = new RequestMessageBody();
            body.Actions = new List<ActionInfo>();
            body.Actions.Add(new ActionInfo() { Id = 1, ActionType = ActionType.Method, Name = "Trace", ObjectPathId = 2, ArgumentInfo = new ArgumentInfo() { Arguments = new object[] { "Hello" } } });
            body.Actions.Add(new ActionInfo() { Id = 2, ActionType = ActionType.Method, Name = "Trace", ArgumentInfo = new ArgumentInfo() { Arguments = new object[] { "Hello" } } });

            string str = Utility.ToJsonString(body);
            this.TestContext.WriteLine("Body={0}", str);
            Assert.IsTrue(true, str);
        }
    }
}
