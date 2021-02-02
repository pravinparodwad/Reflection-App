using System;
using System.Collections.Generic;
using System.Text;
using Attachmate.Reflection.Framework;
using Attachmate.Reflection.Emulation.IbmHosts;
using Attachmate.Reflection.UserInterface;
using Attachmate.Reflection;

namespace ReflectionUtilities
{
    class Program
    {
        public string OpenProfileDocument(string documentFilePath)
        {
            Application app;
            IIbmTerminal terminal;
            IFrame frame;
            IView view;
            string output = string.Empty;
            try
            {
                //Start a visible instance of Reflection or get the instance running at the given channel name
                app = MyReflection.CreateApplication("myWorkspace", true);
                //Create a terminal from the session document file
                //string sessionPath = documentFilePath;
                terminal = (IIbmTerminal)app.CreateControl(documentFilePath);
                //Make the session visible in the workspace
                frame = (IFrame)app.GetObject("Frame");
                view = frame.CreateView(terminal);
                if (view != null)
                {
                    output = "Pass: Successfully opened the profile " + documentFilePath;
                }
                else
                {
                    output = "Fail: Could not find the profile " + documentFilePath;
                }
            }
            catch (Exception e)
            {
                output = "Fail: Error occured while opening profile " + documentFilePath + e.StackTrace;
            }
            return output;
        }
        public IIbmTerminal GetOpenSession(string sessionProfile)
        {
            Application app;
            IFrame frame;
            IView view;
            IIbmTerminal terminal;
            try
            {
                app = MyReflection.CreateApplication();
                frame = (IFrame)app.GetObject("Frame");
                view = frame.GetViewByTitleText(sessionProfile);
                terminal = (IIbmTerminal)view.Control;
            }
            catch (Exception)
            {
                terminal = null;
                //Console.WriteLine(e.StackTrace);
            }
            return terminal;
        }
        public string ConnectSession(string sessionProfile)
        {
            IIbmTerminal terminal;
            ReturnCode returnC;
            string output = string.Empty;
            terminal = GetOpenSession(sessionProfile);
            if (terminal != null)
            {
                returnC = terminal.Connect();
                if (returnC.ToString().Equals("Success"))
                {
                    output = "Pass: Session " + sessionProfile + " is connected successfully.";
                }
                else
                {
                    output = "Fail: Error occured while connecting to session " + sessionProfile + " Status is " + returnC;
                }
            }
            else
            {
                output = "Fail: Not able to find terminal with profile " + sessionProfile;
            }
            return output;
        }
        public string DisconnectSession(string sessionProfile)
        {
            IIbmTerminal terminal;
            string output = string.Empty;
            terminal = GetOpenSession(sessionProfile);
            if (terminal != null)
            {
                terminal.Disconnect();
                output = "Pass: Session " + sessionProfile + " is disconnected successfully.";
            }
            else
            {
                output = "Fail: Not able to find terminal with profile " + sessionProfile;
            }
            return output;
        }
        public bool SetCursorPositionAt(IIbmTerminal terminal, int row, int col)
        {
            ReturnCode returnC;
            returnC = terminal.Screen.MoveCursorTo(row, col);
            if (returnC.ToString().Equals("Success"))
            {
                return true;
            }
            return false;
        }
        public int[] FindCoOrdinates(IIbmTerminal terminal, string positionString)
        {
            //Console.WriteLine("inside FindCoOrdinates methd");
            int[] output = { -1, -1 };
            if (positionString.StartsWith("@Cursor"))
            {
                //Console.WriteLine("inside @Cursor block");
                output[0] = terminal.Screen.CursorRow;
                output[1] = terminal.Screen.CursorColumn;
                //Console.WriteLine("inside @Cursor :" +output[0].ToString()+" ,"+ output[1].ToString());
            }
            else if (positionString.StartsWith("@Position:"))
            {
                positionString = positionString.Replace("@Position:", "");
                output[0] = int.Parse(positionString.Split(',')[0]);
                output[1] = int.Parse(positionString.Split(',')[1]);
            }
            else
            {
                output = null;
            }
            return output;
        }
        public string EnterTextValue(string sessionProfile, string position, string valueToEnter)
        {
            //Console.WriteLine("Inside EnterTextValue method");
            IIbmTerminal terminal;
            bool isCursorMoved = false;
            //int rowNo, colNo;
            string output = string.Empty;
            terminal = GetOpenSession(sessionProfile);
            if (terminal != null)
            {
                //Console.WriteLine("terminal found");
                int[] coOrdinatesArray = FindCoOrdinates(terminal, position);
                if (coOrdinatesArray == null)
                {
                    output = "Fail: Invalid position type provided";
                    return output;
                }
                else if (coOrdinatesArray[0].Equals(string.Empty) && coOrdinatesArray[1].Equals(string.Empty))
                {
                    //Console.WriteLine("Inside @Cursor if block");
                    isCursorMoved = true;
                }
                else if (coOrdinatesArray != null && !coOrdinatesArray[0].Equals(string.Empty) && !coOrdinatesArray[1].Equals(string.Empty))//if(Int32.Parse(coOrdinatesArray[0]) > 0 && Int32.Parse(coOrdinatesArray[1]) > 0)
                {
                    //Console.WriteLine("Inside @Position if block");
                    //rowNo = Int32.Parse(coOrdinatesArray[0]);
                    //colNo = Int32.Parse(coOrdinatesArray[1]);
                    isCursorMoved = SetCursorPositionAt(terminal, coOrdinatesArray[0], coOrdinatesArray[1]);
                }

                if (isCursorMoved)
                {
                    try
                    {
                        terminal.Screen.SendControlKey(ControlKeyCode.DeleteWord);
                        terminal.Screen.SendKeys(valueToEnter);
                    }
                    catch (Exception)
                    {
                        output = "Fail: Error occured while entering the text on specified location " + position;
                    }
                    output = "Pass: Text enterd successfully.";
                }
                else
                {
                    output = "Fail: Error occured while moving cursor to specified location.";
                }
            }
            else
            {
                output = "Fail: Not able to get the terminal window for " + sessionProfile;
            }
            return output;
        }
        public bool SearchStringOnEntireScreen(IIbmTerminal terminal, string textToVerify)
        {
            ScreenPoint point;
            try
            {
                point = terminal.Screen.SearchText(textToVerify, 1, 1, FindOption.Forward);
                //Console.WriteLine(point.ToString());
                string temp = terminal.Screen.GetText(point.Row, point.Column, textToVerify.Length);
                //Console.WriteLine(temp);
                if (temp.Equals(textToVerify))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
        public string VerifyMessage(string sessionProfile, string textToVerify)
        {
            IIbmTerminal terminal;
            bool isMessageFound = false;
            terminal = GetOpenSession(sessionProfile);
            if (terminal != null)
            {
                //Console.WriteLine("Got the session.");
                isMessageFound = SearchStringOnEntireScreen(terminal, textToVerify);
                //Console.WriteLine("returned from search method");
                if (isMessageFound)
                {
                    return "Pass: The specified text is found on session";
                }
                return "Fail: The specified text is not found on session";
            }
            else
            {
                return "Fail: Not able to get the terminal window for " + sessionProfile;
            }

        }
        public string VerifyFieldValue(string sessionProfile, string position, string textToVerify)
        {
            //Console.WriteLine("Inside verifyFieldValue method.");
            IIbmTerminal terminal;
            string textAtGivenPosition = string.Empty;
            string output = string.Empty;
            terminal = GetOpenSession(sessionProfile);
            //Console.WriteLine("Got terminal.");
            int[] coOrdinatesArray = FindCoOrdinates(terminal, position);
            //Console.WriteLine(coOrdinatesArray[0]);
            //Console.WriteLine(coOrdinatesArray[1]);
            try
            {
                if (terminal != null)
                {
                    textAtGivenPosition = terminal.Screen.GetFieldText(coOrdinatesArray[0], coOrdinatesArray[1]);
                    if (textAtGivenPosition.Trim().Equals(textToVerify))
                    {
                        output = "Pass: Specified text is found at given position";
                    }
                    else
                    {
                        output = "Fail: Specified text is not found at given position. Actual text found is:" + textAtGivenPosition;
                    }
                }
                else
                {
                    output = "Fail: Not able to get the terminal window for " + sessionProfile;
                }

            }
            catch (Exception)
            {
                output = "Fail: Error occured while getting field text";
            }

            return output;
        }
        public string SendKey(string sessionProfile, string keyName)
        {
            string output = string.Empty;
            IIbmTerminal terminal;
            ReturnCode returnC = ReturnCode.Error;
            terminal = GetOpenSession(sessionProfile);
            if (terminal != null)
            {
                switch (keyName.ToUpper())
                {
                    case "F1":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F1);
                        break;
                    case "F2":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F2);
                        break;
                    case "F3":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F3);
                        break;
                    case "F4":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F4);
                        break;
                    case "F5":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F5);
                        break;
                    case "F6":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F6);
                        break;
                    case "F7":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F7);
                        break;
                    case "F8":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F8);
                        break;
                    case "F9":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F9);
                        break;
                    case "F10":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F10);
                        break;
                    case "F11":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F11);
                        break;
                    case "F12":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F12);
                        break;
                    case "F13":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F13);
                        break;
                    case "F14":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F14);
                        break;
                    case "F15":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F15);
                        break;
                    case "F16":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F16);
                        break;
                    case "F17":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F17);
                        break;
                    case "F18":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F18);
                        break;
                    case "F19":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F19);
                        break;
                    case "F20":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F20);
                        break;
                    case "F21":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F21);
                        break;
                    case "F22":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F22);
                        break;
                    case "F23":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F23);
                        break;
                    case "F24":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.F24);
                        break;
                    case "ENTER":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Transmit);
                        break;
                    case "TAB":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Tab);
                        break;
                    case "CLEAR":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Clear);
                        break;
                    case "DELETE":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Delete);
                        break;
                    case "DOWN":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Down);
                        break;
                    case "UP":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Up);
                        break;
                    case "LEFT":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Left);
                        break;
                    case "RIGHT":
                        returnC = terminal.Screen.SendControlKey(ControlKeyCode.Right);
                        break;
                }
                if (returnC.Equals(ReturnCode.Success))
                {
                    output = "Pass: Specified key " + keyName + " is sent successfully.";
                }
                else
                {
                    output = "Fail: Not able to send specified key " + keyName;
                }

            }
            else
            {
                output = "Fail: Not able to get the terminal window for " + sessionProfile;
            }
            return output;
        }
        public string GetFieldValue(string sessionProfile, string position)
        {
            //Console.WriteLine("Inside verifyFieldValue method.");
            IIbmTerminal terminal;
            string textAtGivenPosition = string.Empty;
            string output = string.Empty;
            terminal = GetOpenSession(sessionProfile);
            //Console.WriteLine("Got terminal.");
            int[] coOrdinatesArray = FindCoOrdinates(terminal, position);
            //Console.WriteLine(coOrdinatesArray[0]);
            //Console.WriteLine(coOrdinatesArray[1]);
            try
            {
                if (terminal != null)
                {
                    textAtGivenPosition = terminal.Screen.GetFieldText(coOrdinatesArray[0], coOrdinatesArray[1]);
                    if (textAtGivenPosition != null)
                    {
                        output = "Pass: Text at specified position is :" + textAtGivenPosition;
                    }
                    else if (textAtGivenPosition.Equals(string.Empty))
                    {
                        output = "Pass: Text at specified position is :" + textAtGivenPosition;
                    }
                    else
                    {
                        output = "Fail: Not able to fetch text from given position";
                    }

                }
                else
                {
                    output = "Fail: Not able to get the terminal window for " + sessionProfile;
                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                output = "Fail: " + e.ToString();
            }
            catch (Exception)
            {
                output = "Fail: Error occured while getting field text";
            }

            return output;
        }
        public string GetText(string sessionProfile, string position, string length)
        {
            //Console.WriteLine("Inside GetText method.");
            IIbmTerminal terminal;
            string textAtGivenPosition = string.Empty;
            string output = string.Empty;
            int lengthOfTextToGet;
            terminal = GetOpenSession(sessionProfile);
            //Console.WriteLine("Got terminal.");
            int[] coOrdinatesArray = FindCoOrdinates(terminal, position);
            try
            {
                lengthOfTextToGet = int.Parse(length);
            }
            catch (Exception)
            {
                output = "Fail: Provide numeric length";
                return output;
            }
            try
            {
                if (terminal != null)
                {
                    textAtGivenPosition = terminal.Screen.GetText(coOrdinatesArray[0], coOrdinatesArray[1], lengthOfTextToGet);
                    if (textAtGivenPosition != null)
                    {
                        output = "Pass: Text at specified position is :" + textAtGivenPosition;
                    }
                    else if (textAtGivenPosition.Equals(string.Empty))
                    {
                        output = "Pass: Text at specified position is :" + textAtGivenPosition;
                    }
                    else
                    {
                        output = "Fail: Not able to fetch text from given position";
                    }

                }
                else
                {
                    output = "Fail: Not able to get the terminal window for " + sessionProfile;
                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                output = "Fail: " + e.ToString();
            }
            catch (Exception)
            {
                output = "Fail: Error occured while getting field text";
            }

            return output;
        }
        static void Main(string[] args)
        {
            Program reflectionProgram = new Program();
            try
            {
                string result;
                result = reflectionProgram.OpenProfileDocument("C:\\Projects\\Reflection\\Reflection Excercise\\Extra_IO.rd3x");
                //p.GetOpenSession("Extra_IO.rd3x");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.ConnectSession("Extra_IO.rd3x");
                Console.WriteLine(result);
                result = "";
                System.Threading.Thread.Sleep(5000);
                result = reflectionProgram.VerifyMessage("Extra_IO.rd3x", "IOMAINT"); //This will fail
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.VerifyMessage("Extra_IO.rd3x", "Userid");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.EnterTextValue("Extra_IO.rd3x", "@Position:10,32", "IOMAINT");
                Console.WriteLine(result);
                result = "";
                System.Threading.Thread.Sleep(2000);
                result = reflectionProgram.SendKey("Extra_IO.rd3x", "Tab");
                Console.WriteLine(result);
                result = "";
                System.Threading.Thread.Sleep(2000);
                result = reflectionProgram.EnterTextValue("Extra_IO.rd3x", "@Cursor", "I1MAINT");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.VerifyFieldValue("Extra_IO.rd3x", "@Cursor", "IOMAINT");//This will fail.
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.VerifyFieldValue("Extra_IO.rd3x", "@Position:10,32", "IOMAINT");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.SendKey("Extra_IO.rd3x", "Enter");
                Console.WriteLine(result);

                result = "";
                result = reflectionProgram.GetFieldValue("Extra_IO.rd3x", "@Position:16,21");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.GetFieldValue("Extra_IO.rd3x", "@Position:10,4");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.GetText("Extra_IO.rd3x", "@Position:11,1", "80");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.GetText("InvestOne.rd3x", "@Position:11,1", "80");
                Console.WriteLine(result);
                result = "";
                result = reflectionProgram.GetText("InvestOnePravin.rd3x", "@Position:11,1", "80");
                Console.WriteLine(result);
                //result = "";
                //result = reflectionProgram.DisconnectSession("Extra_IO.rd3x");
                //Console.WriteLine(result);
            }
            catch (Exception e)
            {
                e.ToString();
            }
        }
    }
}

//csc ReflectionOperations.cs /r:Attachmate.Reflection.dll,Attachmate.Reflection.Emulation.IbmHosts.dll,Attachmate.Reflection.Emulation.OpenSystems.dll,Attachmate.Reflection.Framework.dll,Attachmate.Reflection.UserControl.IbmHosts.dll,Attachmate.Reflection.UserControl.OpenSystems.dll,Attachmate.Reflection.UserControl.OpenSystemsGraphics.dll,Attachmate.Reflection.UserControl.wpf.IbmHosts.dll,Attachmate.Reflection.UserControl.wpf.OpenSystems.dll,Attachmate.Reflection.UserControl.wpf.OpenSystemsGraphics.dll