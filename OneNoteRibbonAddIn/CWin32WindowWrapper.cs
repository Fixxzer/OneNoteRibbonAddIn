/****************************** Module Header ******************************\
Module Name:  CWin32WindowWrapper.cs
Project:      OneNoteRibbonAddIn
Copyright (c) Microsoft Corporation.

wrapper Win32 HWND handles

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using System;
using System.Windows.Forms;

namespace OneNoteRibbonAddIn
{
    internal class CWin32WindowWrapper : IWin32Window
    {
        private readonly IntPtr _windowHandle;

        public CWin32WindowWrapper(IntPtr windowHandle)
        {
            _windowHandle = windowHandle;
        }

        public IntPtr Handle
        {
            get { return _windowHandle; }
        }
    }
}
