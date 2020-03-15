using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

[
    Guid("B6AF4673-200B-413c-8536-1F778AC14DE1"),
    ProgId("Kerr.Sample.RtdServer"),
    ComVisible(true)
]
public class RtdServer : IRtdServer
{
    private IRTDUpdateEvent m_callback;
    private Timer m_timer;
    private Dictionary<int, string> m_topics;

    public int ServerStart(IRTDUpdateEvent callback)
    {
        m_callback = callback;

        m_timer = new Timer();
        m_timer.Tick += new EventHandler(TimerEventHandler);
        m_timer.Interval = 2000;

        m_topics = new Dictionary<int, string>();

        return 1;
    }

    public void ServerTerminate()
    {
        if (null != m_timer)
        {
            m_timer.Dispose();
            m_timer = null;
        }
    }

    public object ConnectData(int topicId,
                              ref Array strings,
                              ref bool newValues)
    {
        if (1 != strings.Length)
        {
            return "Exactly one parameter is required (e.g. 'hh:mm:ss').";
        }

        string format = strings.GetValue(0).ToString();

        m_topics[topicId] = format;
        m_timer.Start();
        return GetTime(format);
    }

    public void DisconnectData(int topicId)
    {
        m_topics.Remove(topicId);
    }

    public Array RefreshData(ref int topicCount)
    {
        object[,] data = new object[2, m_topics.Count];

        int index = 0;

        foreach (int topicId in m_topics.Keys)
        {
            data[0, index] = topicId;
            data[1, index] = GetTime(m_topics[topicId]);

            ++index;
        }

        topicCount = m_topics.Count;

        m_timer.Start();
        return data;
    }

    public int Heartbeat()
    {
        return 1;
    }

    private void TimerEventHandler(object sender,
                                   EventArgs args)
    {
        m_timer.Stop();
        m_callback.UpdateNotify();
    }

    private static string GetTime(string format)
    {
        return DateTime.Now.ToString(format, CultureInfo.CurrentCulture);
    }
}