using System.Collections.Generic;

/// <summary>
/// ���� ��Ʈ�� ù��° ���� �̸��� ����.
/// ù��° ���� ����(string ���� �ڸ�Ʈ�� �ʿ�X).
/// ���Ŀ� �߰� ����.
/// </summary>
public class DataTestInfo
{
    public int ID;
    public int RANK;
}

public class DataTest : Data<DataTest, DataTestInfo>
{
    private Dictionary<int, DataTestInfo> _dicTest = new Dictionary<int, DataTestInfo>();

    protected override void SetData_Inner()
    {
        for (int i = 0; i < _listNestedData.Count; i++)
        {
            _dicTest.Add(_listNestedData[i].ID, _listNestedData[i]);
        }
    }

    public int GetTestID(int nID)
    {
        return GetChapterData(nID).ID;
    }

    public int GetTestRank(int nID)
    {
        return GetChapterData(nID).RANK;
    }

    private DataTestInfo GetChapterData(int nID)
    {
        return _dicTest.ContainsKey(nID) ? _dicTest[nID] : null;
    }
}
