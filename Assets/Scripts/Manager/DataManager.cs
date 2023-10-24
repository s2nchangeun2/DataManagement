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
        //��ũ��Ʈ�̸��� ������Ʈ �̸� �����ϰ� ���߱�.
        string strPath = Application.dataPath + $"/Data/Json/{typeof(T).Name}.Json";
        string strJson = File.ReadAllText(strPath);

        _listNestedData = JsonConvert.DeserializeObject<List<TT>>(strJson);

        SetData_Inner();
    }

    /// <summary>
    /// �� ������ ���� ��ųʸ��� �߰�.
    /// </summary>
    protected virtual void SetData_Inner() { }
}

/// <summary>
/// ���� ������ ����.
/// </summary>
public class DataManager : SingletonManager<DataManager>
{
    /// <summary>
    /// Type - class.
    /// Data - ������ ���̺�
    /// </summary>      
    private Dictionary<Type, Data> _dicData = new Dictionary<Type, Data>();

    private void Awake()
    {
        //�� ������ �Ŵ��� ��ųʸ��� �߰�.
        //

        //ex.test.
        _dicData.Add(typeof(DataTest), new DataTest());
    }

    public Data GetData(Type type)
    {
        return _dicData.ContainsKey(type) ? _dicData[type] : null;
    }
}