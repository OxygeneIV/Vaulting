using System;
using OpenQA.Selenium;

namespace Framework.Exceptions
{
    [Serializable]
    public class MultipleElementHitsException : WebDriverException
    {
        public MultipleElementHitsException()
        { }
        public MultipleElementHitsException(string message) : base(message) { }
        public MultipleElementHitsException(string message, WebDriverException inner) : base(message, inner) { }

        // A constructor is needed for serialization when an
        // exception propagates from a remoting server to the client. 
        protected MultipleElementHitsException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
        { }
    }
    public class WaitUntilException : Exception
    {
        public WaitUntilException()
        { }
        public WaitUntilException(string message) : base(message) { }
        public WaitUntilException(string message, Exception inner) : base(message, inner) { }

        // A constructor is needed for serialization when an
        // exception propagates from a remoting server to the client. 
        protected WaitUntilException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
        { }
    }
}
