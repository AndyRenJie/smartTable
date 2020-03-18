package com.bin.david.smarttable;

import android.content.Context;
import android.content.Intent;
import android.graphics.Rect;
import android.os.Bundle;
import android.support.design.widget.FloatingActionButton;
import android.support.design.widget.Snackbar;
import android.support.v7.widget.LinearLayoutManager;
import android.support.v7.widget.RecyclerView;
import android.text.TextUtils;
import android.util.Log;
import android.view.View;
import android.support.design.widget.NavigationView;
import android.support.v4.view.GravityCompat;
import android.support.v4.widget.DrawerLayout;
import android.support.v7.app.ActionBarDrawerToggle;
import android.support.v7.app.AppCompatActivity;
import android.support.v7.widget.Toolbar;
import android.view.Menu;
import android.view.MenuItem;
import android.view.inputmethod.InputMethodManager;
import android.widget.Button;
import android.widget.EditText;

import com.bin.david.form.core.SmartTable;
import com.bin.david.form.data.column.Column;
import com.bin.david.form.data.style.FontStyle;
import com.bin.david.form.data.table.TableData;
import com.bin.david.form.utils.DensityUtils;
import com.bin.david.smarttable.adapter.SheetAdapter;
import com.bin.david.smarttable.excel.ExcelCallback;
import com.bin.david.smarttable.excel.IExcel2Table;
import com.bin.david.smarttable.excel.JXLExcel2Table;
import com.chad.library.adapter.base.BaseQuickAdapter;
import com.vincent.filepicker.Constant;
import com.vincent.filepicker.activity.NormalFilePickActivity;
import com.vincent.filepicker.filter.entity.NormalFile;

import java.util.ArrayList;
import java.util.List;

import jxl.Cell;

public class ChoiceExcelActivity extends AppCompatActivity implements NavigationView.OnNavigationItemSelectedListener, ExcelCallback {
    private SmartTable<Cell> table;
    private RecyclerView recyclerView;
    private IExcel2Table<Cell> iExcel2Table;

    private EditText searchInputET;
    private Button searchBtn;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_choice_excel);
        Toolbar toolbar = (Toolbar) findViewById(R.id.toolbar);
        setSupportActionBar(toolbar);

        DrawerLayout drawer = (DrawerLayout) findViewById(R.id.drawer_layout);
        ActionBarDrawerToggle toggle = new ActionBarDrawerToggle(
                this, drawer, toolbar, R.string.navigation_drawer_open, R.string.navigation_drawer_close);
        drawer.addDrawerListener(toggle);
        toggle.syncState();

        searchInputET = (EditText) findViewById(R.id.search_input_et);
        searchBtn = (Button) findViewById(R.id.search_button);

        NavigationView navigationView = (NavigationView) findViewById(R.id.nav_view);
        navigationView.setNavigationItemSelectedListener(this);
        FontStyle.setDefaultTextSize(DensityUtils.sp2px(this, 15));
        recyclerView = (RecyclerView) findViewById(R.id.recyclerView);
        recyclerView.setLayoutManager(new LinearLayoutManager(this, LinearLayoutManager.HORIZONTAL, false));
        table = (SmartTable<Cell>) findViewById(R.id.table);
        iExcel2Table = new JXLExcel2Table();
        iExcel2Table.initTableConfig(this, table);
        iExcel2Table.setCallback(this);
        iExcel2Table.loadSheetList(this, "c.xls");
    }

    @Override
    protected void onDestroy() {
        if (iExcel2Table != null) {
            iExcel2Table.clear();
        }
        iExcel2Table = null;
        super.onDestroy();
    }

    @Override
    public void getSheetListSuc(List<String> sheetNames) {
        recyclerView.setHasFixedSize(true);
        if (sheetNames != null && sheetNames.size() > 0) {
            final SheetAdapter sheetAdapter = new SheetAdapter(sheetNames);
            sheetAdapter.setOnItemClickListener(new BaseQuickAdapter.OnItemClickListener() {
                @Override
                public void onItemClick(BaseQuickAdapter adapter, View view, int position) {
                    sheetAdapter.setSelectPosition(position);
                    iExcel2Table.loadSheetContent(ChoiceExcelActivity.this, position, ChoiceExcelActivity.this);
                }
            });
            recyclerView.setAdapter(sheetAdapter);
            iExcel2Table.loadSheetContent(this, 0, this);
        }
    }

    @Override
    public void getSheelContentSuc() {
        if (table.getTableData() != null) {
            table.getTableData().setOnItemClickListener(new TableData.OnItemClickListener() {
                @Override
                public void onClick(Column column, String value, Object o, int col, int row) {
                    Log.d("Andy.R", "value: " + value + ", object: " + o + ", col: " + col + ", row: " + row);
                }
            });
            searchBtn.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) {
                    //点击搜索完关闭软键盘
                    InputMethodManager inputMethodManager = (InputMethodManager) getSystemService(Context.INPUT_METHOD_SERVICE);
                    if (inputMethodManager != null) {
                        inputMethodManager.hideSoftInputFromWindow(searchInputET.getWindowToken(), InputMethodManager.HIDE_NOT_ALWAYS);
                    }
                    String searchContent = searchInputET.getText().toString().trim();
                    if (TextUtils.isEmpty(searchContent)) {
                        return;
                    }
                    List<Column> columns = table.getTableData().getColumns();
                    for (Column column : columns) {
                        if (column != null) {
                            List<Cell> datas = column.getDatas();
                            if (!datas.isEmpty()) {
                                for (Cell cell : datas) {
                                    if (cell != null && cell.getContents().contains(searchContent)) {
                                        //搜索到的内容所在的列
                                        int targetColumn = cell.getColumn();
                                        //搜索到的内容所在的行
                                        int targetRow = cell.getRow();
                                        //总共多少列
                                        int totalColumn = table.getTableData().getColumns().size();
                                        //总共多少行
                                        int totalRow = table.getTableData().getLineSize();
                                        //如果搜索到的内容所在的列数小于总共列数的一半，代表目标在左边，所以让表格滑动到左边，反之滑动到右边
                                        if (targetColumn < totalColumn / 2) {
                                            table.getMatrixHelper().flingLeft(500);
                                        } else {
                                            table.getMatrixHelper().flingRight(500);
                                        }
                                        //如果搜索到的内容所在的行数小于总共行数的一半，代表目标在上边，所以让表格滑动到上边，反之滑动到下边
                                        if (targetRow < totalRow / 2) {
                                            table.getMatrixHelper().flingTop(500);
                                        } else {
                                            table.getMatrixHelper().flingBottom(500);
                                        }
                                        //绘制当前搜索到的单元格背景，拿到当前单元格的行和列，获取当前单元格的点的位置
                                        int[] pointLocation = table.getProvider().getPointLocation(cell.getRow(), cell.getColumn());
                                        //拿到点的位置生成一个矩形
                                        Rect rect = new Rect(pointLocation[0], pointLocation[1], pointLocation[0], pointLocation[1]);
                                        //设置选中矩形区域
                                        table.getProvider().getOperation().setSelectionRect(cell.getColumn(), cell.getRow(), rect);
                                        //然后刷新
                                        table.invalidate();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            });
        }
    }

    @Override
    public void onBackPressed() {
        DrawerLayout drawer = (DrawerLayout) findViewById(R.id.drawer_layout);
        if (drawer.isDrawerOpen(GravityCompat.START)) {
            drawer.closeDrawer(GravityCompat.START);
        } else {
            super.onBackPressed();
        }
    }

    @SuppressWarnings("StatementWithEmptyBody")
    @Override
    public boolean onNavigationItemSelected(MenuItem item) {
        // Handle navigation view item clicks here.
        int id = item.getItemId();

        if (id == R.id.nav_import) {
            Intent intent4 = new Intent(this, NormalFilePickActivity.class);
            intent4.putExtra(Constant.MAX_NUMBER, 1);
            intent4.putExtra(NormalFilePickActivity.SUFFIX, new String[]{"xls"});
            startActivityForResult(intent4, Constant.REQUEST_CODE_PICK_FILE);
        } else if (id == R.id.nav_export) {

        }
        DrawerLayout drawer = (DrawerLayout) findViewById(R.id.drawer_layout);
        drawer.closeDrawer(GravityCompat.START);
        return true;
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        switch (requestCode) {
            case Constant.REQUEST_CODE_PICK_FILE:
                if (resultCode == RESULT_OK) {
                    ArrayList<NormalFile> list = data.getParcelableArrayListExtra(Constant.RESULT_PICK_FILE);
                    if (list != null && list.size() > 0) {
                        iExcel2Table.setIsAssetsFile(false);
                        iExcel2Table.loadSheetList(this, list.get(0).getPath());
                    }
                }
                break;
        }
    }
}
