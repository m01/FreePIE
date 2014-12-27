using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using FreePIE.Core.Contracts;
using FreePIE.Core.Plugins.Strategies;
using SlimDX.DirectInput;
using SlimDX.RawInput;
//using System.Windows.Forms;  // for Cursor.Position

namespace FreePIE.Core.Plugins
{

    [GlobalType(Type = typeof(MouseGlobal))]
    public class MousePlugin : Plugin
    {
        // Mouse position state variables
        private double deltaXOut;
        private double deltaYOut;
        // and the same for absolute movement mode, see http://msdn.microsoft.com/en-us/library/windows/desktop/ms646273(v=vs.85).aspx for details
        private const uint SCALING_FACTOR = 65535;
        private uint XOut, YOut;
        // set to false for relative (delta) movement mode
        public bool absolute_mode = true;
        private int wheel;
        public const int WheelMax = 120;

        private DirectInput directInputInstance = new DirectInput();
        private Mouse mouseDevice;
        private MouseState currentMouseState;
        private bool leftPressed;
        private bool rightPressed;
        private bool middlePressed;
        private GetPressedStrategy<int> getButtonPressedStrategy;
        private SetPressedStrategy setButtonPressedStrategy;

        public override object CreateGlobal()
        {
            return new MouseGlobal(this);
        }

        public override Action Start()
        {
            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;

            mouseDevice = new Mouse(directInputInstance);
            if (mouseDevice == null)
                throw new Exception("Failed to create mouse device");

            mouseDevice.SetCooperativeLevel(handle, CooperativeLevel.Background | CooperativeLevel.Nonexclusive);
            mouseDevice.Properties.AxisMode = DeviceAxisMode.Relative;   // Get delta values
            mouseDevice.Acquire();

            getButtonPressedStrategy = new GetPressedStrategy<int>(IsButtonDown);
            setButtonPressedStrategy = new SetPressedStrategy(SetButtonDown, SetButtonUp);
          
            OnStarted(this, new EventArgs());
            return null;
        }

        public override void Stop()
        {
            if (mouseDevice != null)
            {
                mouseDevice.Unacquire();
                mouseDevice.Dispose();
                mouseDevice = null;
            }

            if (directInputInstance != null)
            {
                directInputInstance.Dispose();
                directInputInstance = null;
            }
        }
        
        public override string FriendlyName
        {
            get { return "Mouse"; }
        }

        static private MouseKeyIO.MOUSEINPUT MouseInput(int x, int y, uint data, uint t, uint flag)
        {
           var mi = new MouseKeyIO.MOUSEINPUT {dx = x, dy = y, mouseData = data, time = t, dwFlags = flag};
            return mi;
        }

        public override void DoBeforeNextExecute()
        {
            // If a mouse command was given in the script, issue it all at once right here
            if ((int)deltaXOut != 0 || (int)deltaYOut != 0 || (int)XOut != 0 || (int)YOut != 0 || wheel != 0)
            {

                var input = new MouseKeyIO.INPUT[1];
                input[0].type = MouseKeyIO.INPUT_MOUSE;
                if (absolute_mode)
                {
                    input[0].mi = MouseInput((int)XOut, (int)YOut, (uint)wheel, 0, MouseKeyIO.MOUSEEVENTF_MOVE | MouseKeyIO.MOUSEEVENTF_ABSOLUTE | MouseKeyIO.MOUSEEVENTF_WHEEL);
                }
                else
                {
                    input[0].mi = MouseInput((int)deltaXOut, (int)deltaYOut, (uint)wheel, 0, MouseKeyIO.MOUSEEVENTF_MOVE | MouseKeyIO.MOUSEEVENTF_WHEEL);
                }

                MouseKeyIO.SendInput(1, input, Marshal.SizeOf(input[0].GetType()));

                // Reset the mouse values
                if ((int)deltaXOut != 0)
                {
                    deltaXOut = deltaXOut - (int)deltaXOut;
                }
                if ((int)deltaYOut != 0)
                {
                    deltaYOut = deltaYOut - (int)deltaYOut;
                }

                if ((int)XOut != 0)
                {
                    XOut = 0;
                }
                if ((int)YOut != 0)
                {
                    YOut = 0;
                }

                wheel = 0;
            }

            currentMouseState = null;  // flush the mouse state

            setButtonPressedStrategy.Do();
        }

        /*
         * Input: X coordinate as a normalised value (0..1.0)
         * Currently this is read-only until:
         * - a way to set the X coordinate using pixel numbers is found, or
         * - a way to normalise the pixel coordinates returned by Cursor.Position.{X,Y} is found.
         */
        public double X
        {
            set
            {
                if (0.0 <= value && value <= 1.0)
                {
                    XOut = (uint)(value * SCALING_FACTOR);
                }
                else
                {
                    throw new ArgumentOutOfRangeException("X", "X needs to be a normalised screen coordinate between 0 and 1.0.");
                }
            }
            //get { return Cursor.Position.X; }
        }

        /*
         * Input: Y coordinate as a normalised value (0..1.0)
         * Currently this is read-only until:
         * - a way to set the X coordinate using pixel numbers is found, or
         * - a way to normalise the pixel coordinates returned by Cursor.Position.{X,Y} is found.
         */
        public double Y
        {
            set 
            {
                if (0.0 <= value && value <= 1.0)
                {
                    YOut = (uint)(value * SCALING_FACTOR); 
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Y", "Y needs to be a normalised screen coordinate between 0 and 1.0.");
                }
            }
            //get { return Cursor.Position.Y; }
        }

        public double DeltaX
        {
            set { deltaXOut = deltaXOut + value; }
            get { return CurrentMouseState.X; }
        }

        public double DeltaY
        {
            set
            {
                deltaYOut = deltaYOut + value;
            }

            get { return CurrentMouseState.Y; }
        }

        public int Wheel
        {
            get { return CurrentMouseState.Z; }
            set { wheel = value; }
            
        }

        private MouseState CurrentMouseState
        {
            get
            {
                if (currentMouseState == null)
                    currentMouseState = mouseDevice.GetCurrentState();

                return currentMouseState;
            }
        }

        public bool IsButtonDown(int index)
        {
            return CurrentMouseState.IsPressed(index);
        }

        public bool IsButtonPressed(int button)
        {
            return getButtonPressedStrategy.IsPressed(button);
        }

        private void SetButtonDown(int button)
        {
            SetButtonPressed(button, true);    
        }

        private void SetButtonUp(int button)
        {
            SetButtonPressed(button, false);
        }

        public void SetButtonPressed(int index, bool pressed)
        {
            uint btn_flag = 0;
            if (index == 0)
            {
               if (pressed)
               {
                  if (!leftPressed)
                     btn_flag = MouseKeyIO.MOUSEEVENTF_LEFTDOWN;
               }
               else
               {
                  if (leftPressed)
                     btn_flag = MouseKeyIO.MOUSEEVENTF_LEFTUP;
               }
               leftPressed = pressed;
            }
            else if (index == 1)
            {
               if (pressed)
               {
                  if (!rightPressed)
                     btn_flag = MouseKeyIO.MOUSEEVENTF_RIGHTDOWN;
               }
               else
               {
                  if (rightPressed)
                     btn_flag = MouseKeyIO.MOUSEEVENTF_RIGHTUP;
               }
               rightPressed = pressed;
            }
            else
            {
               if (pressed)
               {
                  if (!middlePressed)
                     btn_flag = MouseKeyIO.MOUSEEVENTF_MIDDLEDOWN;
               }
               else
               {
                  if (middlePressed)
                     btn_flag = MouseKeyIO.MOUSEEVENTF_MIDDLEUP;
               }
               middlePressed = pressed;
            }
           
            if (btn_flag != 0) {
               var input = new MouseKeyIO.INPUT[1];
               input[0].type = MouseKeyIO.INPUT_MOUSE;
               input[0].mi = MouseInput(0, 0, 0, 0, btn_flag);
            
               MouseKeyIO.SendInput(1, input, Marshal.SizeOf(input[0].GetType()));
            }
        }

        public void PressAndRelease(int button)
        {
            setButtonPressedStrategy.Add(button);
        }
    }

    [Global(Name = "mouse")]
    public class MouseGlobal : UpdateblePluginGlobal<MousePlugin>
    {
        public MouseGlobal(MousePlugin plugin) : base(plugin) { }

        public int wheelMax
        {
            get { return MousePlugin.WheelMax; }
        }

        public double deltaX
        {
            get { return plugin.DeltaX; }
            set { plugin.DeltaX = value; }
        }

        public double deltaY
        {
            get { return plugin.DeltaY; }
            set { plugin.DeltaY = value; }
        }

        /*
         * between 0..1.0
         */
        public double X
        {
            //get { return plugin.X; }  // see comment in the class above
            set { plugin.X = value; }
        }

        /*
         * between 0..1.0
         */
        public double Y
        {
            //get { return plugin.Y; }  // see comment in the class above
            set { plugin.Y = value; }
        }

        /*
         * Set this to True to use X, Y and set absolute positions, or set it to False (default)
         * to use deltaX, deltaY.
         */
        public bool absoluteMode
        {
            get { return plugin.absolute_mode; }
            set { plugin.absolute_mode = value; }
        }

        public int wheel
        {
            get { return plugin.Wheel; }
            set { plugin.Wheel = value; }
        }

        public bool wheelUp
        {
            get { return plugin.Wheel == wheelMax; }
            set { plugin.Wheel = value ? wheelMax : 0; }
        }

        public bool wheelDown
        {
            get { return plugin.Wheel == -wheelMax; }
            set { plugin.Wheel = value ? -wheelMax : 0; }
        }

        public bool leftButton
        {
            get { return plugin.IsButtonDown(0); }
            set { plugin.SetButtonPressed(0, value); }
        }

        public bool middleButton
        {
            get { return plugin.IsButtonDown(2); }
            set { plugin.SetButtonPressed(2, value); }
        }

        public bool rightButton
        {
            get { return plugin.IsButtonDown(1); }
            set { plugin.SetButtonPressed(1, value); }
        }

        public bool getButton(int button)
        {
            return plugin.IsButtonDown(button);
        }

        public void setButton(int button, bool pressed)
        {
            plugin.SetButtonPressed(button, pressed);
        }

        public bool getPressed(int button)
        {
            return plugin.IsButtonPressed(button);
        }

        public void setPressed(int button)
        {
            plugin.PressAndRelease(button);
        }
    }
}
