using System;

/***
 * CheckeCallbackArgv.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.07
 */
namespace xlsx2string
{
    public class CheckeCallbackArgv
    {
        public Action<int> OnProgressChanged
        {
            get;
            set;
        }

        public Action<string> OnRunChanged
        {
            get;
            set;
        }
    }
}
