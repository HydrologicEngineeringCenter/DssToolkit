﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Usgs
{
  internal class TextFile
  {
    string[] data;
    int currentIndex = 0;
    public TextFile(string fileName)
    {
      data = File.ReadAllLines(fileName);
    }

    public string this[int index]
    {
      get { return data[index]; }
    }

    public int FindNextBeginningWith(string text,int startIndex)
    {

      for (int i = startIndex; i < data.Length; i++)
      {
        if (data[i].StartsWith(text))
          return i;
      }

      return -1;

    }

    /// <summary>
    /// Find the next matching value in the file, composed of lines like:
    /// key=value
    /// </summary>
    /// <param name="text">key</param>
    /// <returns></returns>
    internal string GetNext(string text)
    {
      int idx = FindNextBeginningWith(text + "=", currentIndex);
      if (idx < 0)
      {
        return "";
      }

      var s = data[idx];
      var i2 = s.IndexOf("=");
      currentIndex = idx + 1;
      return s.Substring(i2 + 1);
    }

  }
}



