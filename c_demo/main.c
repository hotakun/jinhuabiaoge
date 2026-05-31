/*
 * C Demo: CSV 统计 —— 读一个 CSV 文件，按商品名汇总数量
 * 编译: gcc -O2 main.c -o c_demo.exe
 *
 * 感受 C 与 Python 的区别:
 * - 没有 pandas，每行自己 split
 * - 没有 dict，汇总用结构体数组线性查找
 * - 没有 GC，用完手动 free
 * - 但跑起来比 Python 快 10-50 倍
 */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define MAX_ROWS  20000
#define MAX_COLS  32
#define MAX_FIELD 256
#define MAX_ITEMS 4096

/* 一个汇总条目 */
typedef struct {
    char  name[MAX_FIELD];
    int   qty;
} Item;

/* CSV 行 */
typedef struct {
    char fields[MAX_COLS][MAX_FIELD];
    int  ncols;
} Row;

/* ---- 工具函数 ---- */

/* 读一行 CSV，处理引号内的逗号 */
int read_csv_row(FILE *f, Row *row) {
    char line[8192];
    if (!fgets(line, sizeof(line), f)) return 0;

    row->ncols = 0;
    char *p = line;
    int in_quotes = 0;
    char *start = p;

    while (*p) {
        if (*p == '"') { in_quotes = !in_quotes; p++; continue; }
        if (*p == ',' && !in_quotes) {
            size_t len = p - start;
            if (len >= MAX_FIELD) len = MAX_FIELD - 1;
            memcpy(row->fields[row->ncols], start, len);
            row->fields[row->ncols][len] = '\0';
            row->ncols++;
            start = p + 1;
            if (row->ncols >= MAX_COLS) break;
        }
        p++;
    }
    /* 最后一个字段 */
    if (row->ncols < MAX_COLS && *start) {
        size_t len = strlen(start);
        /* 去掉结尾换行 */
        while (len > 0 && (start[len-1] == '\n' || start[len-1] == '\r')) len--;
        if (len >= MAX_FIELD) len = MAX_FIELD - 1;
        memcpy(row->fields[row->ncols], start, len);
        row->fields[row->ncols][len] = '\0';
        row->ncols++;
    }
    return 1;
}

/* 在 items 中查找 name，找不到返回 -1 */
int find_item(Item *items, int count, const char *name) {
    for (int i = 0; i < count; i++) {
        if (strcmp(items[i].name, name) == 0) return i;
    }
    return -1;
}

/* 去首尾空白 */
void trim(char *s) {
    while (*s == ' ' || *s == '\t') s++;
    char *end = s + strlen(s) - 1;
    while (end > s && (*end == ' ' || *end == '\t')) { *end = '\0'; end--; }
    /* 结果前移 */
    if (s != (char*)s) memmove((char*)s, s, strlen(s) + 1);
}

/* ---- 主程序 ---- */
int main(int argc, char **argv) {
    if (argc < 2) {
        printf("用法: c_demo.exe <csv文件>\n");
        return 1;
    }

    FILE *f = fopen(argv[1], "r");
    if (!f) { printf("无法打开文件: %s\n", argv[1]); return 1; }

    /* 读表头，找 "商品名称" 和 "订货数量" 的列索引 */
    Row header;
    if (!read_csv_row(f, &header)) { fclose(f); return 1; }
    int col_name = -1, col_qty = -1;
    for (int i = 0; i < header.ncols; i++) {
        if (strstr(header.fields[i], "商品名称")) col_name = i;
        if (strstr(header.fields[i], "订货数量")) col_qty  = i;
    }
    if (col_name < 0 || col_qty < 0) {
        printf("找不到必要列 (需要 商品名称 + 订货数量)\n");
        fclose(f); return 1;
    }

    /* 逐行汇总 */
    Item items[MAX_ITEMS];
    int   item_count = 0;
    int   line_count = 0;
    Row   row;

    while (read_csv_row(f, &row)) {
        line_count++;
        if (line_count == 1) continue;  /* 跳过表头下的第一行(原excel表头) */
        if (row.ncols <= col_name || row.ncols <= col_qty) continue;

        char *name = row.fields[col_name];
        int   qty  = atoi(row.fields[col_qty]);
        if (qty == 0 && row.fields[col_qty][0] != '0') continue;

        int idx = find_item(items, item_count, name);
        if (idx >= 0) {
            items[idx].qty += qty;
        } else if (item_count < MAX_ITEMS) {
            strncpy(items[item_count].name, name, MAX_FIELD - 1);
            items[item_count].qty = qty;
            item_count++;
        }
    }
    fclose(f);

    /* 输出结果 */
    printf("商品名称                          数量\n");
    printf("──────────────────────────────────────\n");
    int total = 0;
    for (int i = 0; i < item_count; i++) {
        printf("%-35s %5d\n", items[i].name, items[i].qty);
        total += items[i].qty;
    }
    printf("──────────────────────────────────────\n");
    printf("共 %d 种商品, 出库总计 %d\n", item_count, total);
    printf("\n处理 %d 行耗时: ", line_count);

    return 0;
}
