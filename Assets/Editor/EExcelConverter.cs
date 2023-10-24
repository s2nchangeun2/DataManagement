using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using UnityEditor;
using UnityEngine;
using Newtonsoft.Json;
using Excel;
using Ookii.Dialogs;
using System;

public class ExcelConverter : EditorWindow
{
    public static string strDataPath;

    private int _nIndexFormat = 0;
    private int _nIndexEncoding = 0;

    private string[] _strArrFormat = new string[] { "JSON", "XML" };
    private string[] _strArrEncoding = new string[] { "UTF-8", "GB2312" };

    private bool _bConvertString = false;

    private string _strInputPath = null;
    private string _strOutputPath = null;
    private Stream _stream = null;

    private GUIStyle _gUIStyle = null;

    [UnityEditor.MenuItem("Tools/ExcelConverter")]
    public static void ShowExcelConverter()
    {
        Init();
    }

    private static void Init()
    {
        GetWindow(typeof(ExcelConverter), true, "ExcelConverter");

        strDataPath = UnityEngine.Application.dataPath;
        strDataPath = strDataPath.Substring(0, strDataPath.LastIndexOf("/"));
    }

    private void OnEnable()
    {
        _gUIStyle = new GUIStyle();
        _gUIStyle.normal.background = EditorGUIUtility.Load("node0") as Texture2D;
        _gUIStyle.padding = new RectOffset(30, 30, 30, 30);
        _gUIStyle.border = new RectOffset(10, 10, 10, 10);
    }

    private void OnGUI()
    {
        DrawOptions();
    }

    private void DrawOptions()
    {
        GUILayout.BeginArea(new Rect(10, 10, UnityEngine.Screen.width * 0.8f, 150), _gUIStyle);
        EditorGUILayout.LabelField("변환 형식:");
        _nIndexFormat = EditorGUILayout.Popup(_nIndexFormat, _strArrFormat, GUILayout.Width(UnityEngine.Screen.width * 0.5f));
        _bConvertString = GUI.Toggle(new Rect(10, 10, 100, 20), _bConvertString, "텍스트 변환");
        GUILayout.Space(20);
        EditorGUILayout.LabelField("인코딩 형식:");
        _nIndexEncoding = EditorGUILayout.Popup(_nIndexEncoding, _strArrEncoding, GUILayout.Width(UnityEngine.Screen.width * 0.5f));
        GUILayout.EndArea();

        GUILayout.BeginArea(new Rect(10, 150, UnityEngine.Screen.width * 0.8f, 150), _gUIStyle);
        if (GUILayout.Button("Excel 파일 가져오기"))
        {
            _strInputPath = ShowFileDialog();

            if (!string.IsNullOrEmpty(_strInputPath))
                Debug.Log("Excel File 경로: " + _strInputPath);
        }
        _strInputPath = EditorGUILayout.TextField("Excel File 경로: ", _strInputPath);

        GUILayout.Space(10);
        if (GUILayout.Button("Json 파일 내보내기"))
        {
            _strOutputPath = ShowFolderDialog();

            if (!string.IsNullOrEmpty(_strOutputPath))
                Debug.Log("Json Folder 경로: " + _strOutputPath);
        }
        _strOutputPath = EditorGUILayout.TextField("Json Folder 경로: ", _strOutputPath);
        GUILayout.EndArea();

        GUILayout.BeginArea(new Rect(10, 290, UnityEngine.Screen.width * 0.8f, 80), _gUIStyle);
        if (GUILayout.Button("Convert"))
        {
            Convert();
            Close();
        }
        GUILayout.EndArea();
    }

    private string ShowFileDialog()
    {
        VistaOpenFileDialog VistaOpenFileDialog = new VistaOpenFileDialog();

        if (VistaOpenFileDialog.ShowDialog() == DialogResult.OK)
            return VistaOpenFileDialog.FileName;

        _stream.Close();
        return null;
    }

    private string ShowFolderDialog()
    {
        VistaFolderBrowserDialog vistaFolderBrowserDialog = new VistaFolderBrowserDialog();

        if (vistaFolderBrowserDialog.ShowDialog() == DialogResult.OK)
            return vistaFolderBrowserDialog.SelectedPath;

        return null;
    }

    private void Convert()
    {
        string strInputPath = _strInputPath;
        string strOutputPath = string.Empty;

        ExcelConvert excel = new ExcelConvert(strInputPath);
        Encoding encoding = null;

        switch (_nIndexEncoding)
        {
            case 0:
                encoding = Encoding.GetEncoding("utf-8");
                break;
            case 1:
                encoding = Encoding.GetEncoding("gb2312");
                break;
            default:
                break;
        }

        string strFileName = _strOutputPath + "/" + Path.GetFileName(strInputPath);
        switch (_nIndexFormat)
        {
            case 0:
                strOutputPath = strFileName.Replace(".xlsx", ".json");
                excel.ConvertToJson(_strOutputPath, strOutputPath, encoding, _bConvertString);
                break;
            case 1:
                strOutputPath = strFileName.Replace(".xlsx", ".xml");
                excel.ConvertToXml(_strOutputPath, strOutputPath);
                break;
        }

        AssetDatabase.Refresh();
    }
}


public class ExcelConvert
{
    private DataSet _dataSet;

    public ExcelConvert(string strPathExcelFile)
    {
        FileStream mStream = File.Open(strPathExcelFile, FileMode.Open, FileAccess.Read);
        _dataSet = ExcelReaderFactory.CreateOpenXmlReader(mStream).AsDataSet();
    }

    public void ConvertToJson(string strDirectory, string strPathJsonFile, Encoding encoding, bool bConvertString)
    {
        if (_dataSet.Tables.Count < 1)
            return;

        DataTable dataTable = _dataSet.Tables[0];

        if (dataTable.Rows.Count < 1)
            return;

        int nRowCount = dataTable.Rows.Count;
        int nColumnCount = dataTable.Columns.Count;

        List<Dictionary<string, object>> dicExcelData = new List<Dictionary<string, object>>();
        for (int i = 1; i < nRowCount; i++)
        {
            Dictionary<string, object> dicByRow = new Dictionary<string, object>();
            for (int j = 1; j < nColumnCount; j++)
            {
                //첫번째 행은 보통 속성의 이름(ex.ID, Rank, Price..)
                string strFieldName = dataTable.Rows[0][j].ToString();

                string strValue = dataTable.Rows[i][j].ToString();
                if (strValue.Contains(","))
                {
                    string[] strArray = strValue.Split(",");
                    int[] nValues = new int[strArray.Length];
                    for (int k = 0; k < strArray.Length; k++)
                    {
                        nValues[k] = int.Parse(strArray[k]);
                    }
                    dicByRow[strFieldName] = nValues;
                }
                else
                {
                    dicByRow[strFieldName] = bConvertString ? strValue : Convert.ToInt32(strValue);
                }
            }

            dicExcelData.Add(dicByRow);
        }

        string strJson = JsonConvert.SerializeObject(dicExcelData, Newtonsoft.Json.Formatting.Indented);

        Directory.CreateDirectory(strDirectory);
        using (FileStream fileStream = new FileStream(strPathJsonFile, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, encoding))
            {
                textWriter.Write(strJson);
            }
        }
    }

    public void ConvertToXml(string strDirectory, string strPathXMLFile)
    {
        if (_dataSet.Tables.Count < 1)
            return;

        DataTable mSheet = _dataSet.Tables[0];

        if (mSheet.Rows.Count < 1)
            return;

        int rowCount = mSheet.Rows.Count;
        int colCount = mSheet.Columns.Count;

        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        stringBuilder.Append("\r\n");
        stringBuilder.Append("<Table>");
        stringBuilder.Append("\r\n");

        for (int i = 1; i < rowCount; i++)
        {
            stringBuilder.Append("  <Row>");
            stringBuilder.Append("\r\n");
            for (int j = 1; j < colCount; j++)
            {
                stringBuilder.Append("   <" + mSheet.Rows[0][j].ToString() + ">");
                stringBuilder.Append(mSheet.Rows[i][j].ToString());
                stringBuilder.Append("</" + mSheet.Rows[0][j].ToString() + ">");
                stringBuilder.Append("\r\n");
            }
            stringBuilder.Append("  </Row>");
            stringBuilder.Append("\r\n");
        }

        stringBuilder.Append("</Table>");
        Directory.CreateDirectory(strDirectory);
        using (FileStream fileStream = new FileStream(strPathXMLFile, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.GetEncoding("utf-8")))
            {
                textWriter.Write(stringBuilder.ToString());
            }
        }
    }
}
