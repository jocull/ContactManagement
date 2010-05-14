using System;
using System.Collections;
using System.Windows.Forms;

public class clsListViewComparer : IComparer
{
    private int mColumnToCompare;
    private bool mSortOrder;

    public clsListViewComparer(int column, bool sortOrder)
    {
        mColumnToCompare = column;
        mSortOrder = sortOrder;
    }

    public int Compare(object object1, object object2)
    {
        ListViewItem item1;
        ListViewItem item2;
        string text1;
        string text2;

        int returnValue;

        item1 = (ListViewItem)object1;
        item2 = (ListViewItem)object2;

        text1 = item1.SubItems[mColumnToCompare].Text;
        text2 = item2.SubItems[mColumnToCompare].Text;

        returnValue = String.Compare(
            item1.SubItems[mColumnToCompare].Text,
            item2.SubItems[mColumnToCompare].Text, true);


        //if mSortOrder is true, then the program will sort decending, otherwise sort ascending.
        if (mSortOrder == true)
        {
            return -returnValue;
        }
        else
        {
            return returnValue;
        }

    }
}