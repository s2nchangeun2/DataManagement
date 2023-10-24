using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

public abstract class Data
{
    public Data()
    {
        SetData();
    }

    protected abstract void SetData();
}

public class Data<T, TT> : Data
{
    protected List<TT> _listNestedData = null;

    protected override void SetData()
    {
        //스크립트이름과 엑셀시트 이름 동일하게 맞추기.
        string strPath = Application.dataPath + $"/Data/Json/{typeof(T).Name}.Json";
        string strJson = File.ReadAllText(strPath);

        _listNestedData = JsonConvert.DeserializeObject<List<TT>>(strJson);

        SetData_Inner();
    }

    /// <summary>
    /// 각 데이터 개별 딕셔너리에 추가.
    /// </summary>
    protected virtual void SetData_Inner() { }
}

/// <summary>
/// 엑셀 데이터 관리.
/// </summary>
public class DataManager : SingletonManager<DataManager>
{
    /// <summary>
    /// Type - class.
    /// Data - 각각의 테이블
    /// </summary>      
    private Dictionary<Type, Data> _dicData = new Dictionary<Type, Data>();

    private void Awake()
    {
        //각 데이터 매니저 딕셔너리에 추가.
        //

        //ex.test.
        _dicData.Add(typeof(DataTest), new DataTest());
    }

    public Data GetData(Type type)
    {
        return _dicData.ContainsKey(type) ? _dicData[type] : null;
    }
}