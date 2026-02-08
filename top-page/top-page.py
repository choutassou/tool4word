#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
top-page.py - DOCXファイルのヘッダー、フッター、表紙をテンプレートに従って変更するツール

使用例:
    python top-page.py "sample/21GXP-D-001 ソフトウェア要求仕様書.docx" --DocCode D-001 --DocName ソフトウェア要求仕様書 --Version 1.0
"""

import argparse
import sys
from pathlib import Path
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Mm
from lxml import etree


class DocxFormatter:
    """DOCXファイルをテンプレートに従ってフォーマットするクラス"""

    def __init__(self, template_path, variables):
        """
        Args:
            template_path: テンプレートファイルのパス
            variables: 置換する変数の辞書 {'DocCode': 'D-001', ...}
        """
        self.template_doc = Document(template_path)
        self.variables = variables

    def format(self, source_path, output_path):
        """ソースファイルをフォーマットして出力する

        Args:
            source_path: 変換元ファイルのパス
            output_path: 出力先ファイルのパス
        """
        self.target_doc = Document(source_path)

        # 1. ヘッダーをテンプレートからコピー
        self._copy_headers()

        # 2. フッターをテンプレートからコピー
        self._copy_footers()

        # 3. 表紙をテンプレートで置き換え
        self._replace_cover_page()

        # 出力ディレクトリを作成
        output_dir = output_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)

        # 保存
        self.target_doc.save(output_path)
        print(f'変換完了: {output_path}')

    def _replace_variables_in_element(self, element):
        """XML要素全体のテキストを結合して変数置換し、再分配する

        変数が複数のw:r（run）に分割されている場合に対応
        """
        for para in element.iter(qn('w:p')):
            runs = list(para.iter(qn('w:r')))
            if not runs:
                continue

            # 全テキストを結合
            full_text = ''
            text_elements = []
            for run in runs:
                for t in run.iter(qn('w:t')):
                    if t.text:
                        full_text += t.text
                        text_elements.append(t)

            # 変数が含まれているか確認
            has_variable = any(f'${var}' in full_text for var in self.variables.keys())
            if not has_variable:
                continue

            # 変数を置換
            new_text = full_text
            for var_name, var_value in self.variables.items():
                new_text = new_text.replace(f'${var_name}', var_value)

            # 最初のテキスト要素に全テキストを設定し、残りをクリア
            if text_elements:
                text_elements[0].text = new_text
                text_elements[0].set(qn('xml:space'), 'preserve')
                for t in text_elements[1:]:
                    t.text = ''

    def _remove_tab_before_販売名(self, element):
        """「販売名」の前のTAB文字を削除する（Wordバグ対応）"""
        for para in element.iter(qn('w:p')):
            runs = list(para.iter(qn('w:r')))
            if not runs:
                continue

            # 全テキストを結合して「販売名」が含まれるか確認
            full_text = ''
            for run in runs:
                for t in run.iter(qn('w:t')):
                    if t.text:
                        full_text += t.text

            if '販売名' not in full_text:
                continue

            # 各run内のタブ要素を探して削除
            for i, run in enumerate(runs):
                # このrunに「販売名」が含まれているか確認
                run_has_販売名 = False
                for t in run.iter(qn('w:t')):
                    if t.text and '販売名' in t.text:
                        run_has_販売名 = True
                        break

                if run_has_販売名:
                    # このrun内のタブ要素を削除
                    tabs_to_remove = list(run.iter(qn('w:tab')))
                    for tab in tabs_to_remove:
                        tab.getparent().remove(tab)

                    # 前のrunにタブがあれば削除
                    if i > 0:
                        prev_run = runs[i - 1]
                        prev_tabs = list(prev_run.iter(qn('w:tab')))
                        for tab in prev_tabs:
                            tab.getparent().remove(tab)

    def _copy_header_element(self, source_header, target_header):
        """ヘッダー要素をコピーする"""
        # ターゲットヘッダーをクリア
        for child in list(target_header):
            target_header.remove(child)

        # ソースヘッダーの子要素をコピー
        for child in source_header:
            new_child = deepcopy(child)
            target_header.append(new_child)

        # 変数を置換
        self._replace_variables_in_element(target_header)

        # Wordバグ対応：「販売名」の前のTAB文字を削除
        self._remove_tab_before_販売名(target_header)

    def _copy_headers(self):
        """テンプレートのヘッダーをコピーする

        表紙と表紙以外のヘッダーは別設定として扱う
        """
        for i, target_section in enumerate(self.target_doc.sections):
            template_section = self.template_doc.sections[min(i, len(self.template_doc.sections) - 1)]

            # First page header（表紙用）
            if template_section.first_page_header is not None:
                # different_first_page_header_footerを有効にする
                target_section.different_first_page_header_footer = True

                source_first_header = template_section.first_page_header._element
                target_first_header = target_section.first_page_header._element
                self._copy_header_element(source_first_header, target_first_header)

            # 通常のヘッダー（表紙以外）
            source_header = template_section.header._element
            target_header = target_section.header._element
            self._copy_header_element(source_header, target_header)

    def _create_page_field_footer(self, footer_element):
        """PAGEフィールドのみを含むフッターを作成する

        ページ番号書式は -1- 形式（w:pgNumTypeで設定）
        """
        # フッターをクリア
        for child in list(footer_element):
            footer_element.remove(child)

        # 段落を作成
        p = OxmlElement('w:p')

        # 段落プロパティ（中央揃え）
        pPr = OxmlElement('w:pPr')
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center')
        pPr.append(jc)
        p.append(pPr)

        # runを作成
        r = OxmlElement('w:r')

        # PAGEフィールドを作成
        fldChar_begin = OxmlElement('w:fldChar')
        fldChar_begin.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = ' PAGE '

        fldChar_separate = OxmlElement('w:fldChar')
        fldChar_separate.set(qn('w:fldCharType'), 'separate')

        fldChar_end = OxmlElement('w:fldChar')
        fldChar_end.set(qn('w:fldCharType'), 'end')

        r.append(fldChar_begin)
        r2 = OxmlElement('w:r')
        r2.append(instrText)
        r3 = OxmlElement('w:r')
        r3.append(fldChar_separate)
        r4 = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = '1'  # プレースホルダー
        r4.append(t)
        r5 = OxmlElement('w:r')
        r5.append(fldChar_end)

        p.append(r)
        p.append(r2)
        p.append(r3)
        p.append(r4)
        p.append(r5)

        footer_element.append(p)

    def _set_page_number_format(self, section):
        """セクションのページ番号書式を -1- 形式に設定し、開始番号を0にする"""
        sectPr = section._sectPr

        # 既存のpgNumTypeを探すか作成
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is None:
            pgNumType = OxmlElement('w:pgNumType')
            sectPr.append(pgNumType)

        # ページ番号書式を numberInDash (-1-) に設定
        pgNumType.set(qn('w:fmt'), 'numberInDash')
        # 開始ページ番号を0に設定
        pgNumType.set(qn('w:start'), '0')

    def _copy_footers(self):
        """フッターを設定する

        表紙はフッターを表示しない、他はPAGEフィールドのみで -1- 形式
        """
        for i, target_section in enumerate(self.target_doc.sections):
            # different_first_page_header_footerを有効にする（表紙はフッター非表示のため）
            target_section.different_first_page_header_footer = True

            # ページ番号書式を -1- 形式に設定
            self._set_page_number_format(target_section)

            # First page footer（表紙用）- 空にする
            target_first_footer = target_section.first_page_footer._element
            for child in list(target_first_footer):
                target_first_footer.remove(child)

            # 通常のフッター（PAGEフィールドのみ）
            target_footer = target_section.footer._element
            self._create_page_field_footer(target_footer)

    def _get_cover_page_elements(self, doc):
        """表紙ページの要素（最初のページブレイクまで）を取得する

        Returns:
            list: 表紙の要素リスト
        """
        body = doc._body._body
        elements = []

        for elem in body:
            if elem.tag == qn('w:p'):
                xml_str = etree.tostring(elem, encoding='unicode')
                # ページブレイクが含まれている場合、この段落まで含めて終了
                if 'w:br' in xml_str and 'w:type="page"' in xml_str:
                    elements.append(elem)
                    break
                elements.append(elem)
            elif elem.tag == qn('w:tbl'):
                elements.append(elem)
            elif elem.tag == qn('w:sectPr'):
                break
            else:
                elements.append(elem)

        return elements

    def _set_cell_width(self, tc, width_twips):
        """セルの幅を設定する"""
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)

        tcW = tcPr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW')
            tcPr.insert(0, tcW)

        tcW.set(qn('w:w'), str(width_twips))
        tcW.set(qn('w:type'), 'dxa')

    def _set_table_column_widths(self, table_element, col1_mm, col2_mm):
        """テーブルの1列目と2列目の幅を設定する

        Args:
            table_element: テーブルのXML要素
            col1_mm: 1列目の幅（ミリメートル）
            col2_mm: 2列目の幅（ミリメートル）
        """
        col1_twips = int(col1_mm * 56.7)  # 1mm ≈ 56.7 twips
        col2_twips = int(col2_mm * 56.7)
        total_twips = col1_twips + col2_twips

        # テーブルプロパティを設定（レイアウトを固定に）
        tblPr = table_element.find(qn('w:tblPr'))
        if tblPr is not None:
            # tblLayoutをfixedに設定（自動調整を無効化）
            tblLayout = tblPr.find(qn('w:tblLayout'))
            if tblLayout is None:
                tblLayout = OxmlElement('w:tblLayout')
                tblPr.append(tblLayout)
            tblLayout.set(qn('w:type'), 'fixed')

            # tblWを固定幅に設定
            tblW = tblPr.find(qn('w:tblW'))
            if tblW is None:
                tblW = OxmlElement('w:tblW')
                tblPr.append(tblW)
            tblW.set(qn('w:w'), str(total_twips))
            tblW.set(qn('w:type'), 'dxa')

        # テーブルの直接の子要素であるtr（行）を取得
        for tr in table_element.findall(qn('w:tr')):
            # 行の直接の子要素であるtc（セル）を取得
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) >= 1:
                self._set_cell_width(tcs[0], col1_twips)
            if len(tcs) >= 2:
                self._set_cell_width(tcs[1], col2_twips)

        # テーブルグリッドの列幅も設定
        tblGrid = table_element.find(qn('w:tblGrid'))
        if tblGrid is not None:
            gridCols = tblGrid.findall(qn('w:gridCol'))
            if len(gridCols) >= 1:
                gridCols[0].set(qn('w:w'), str(col1_twips))
            if len(gridCols) >= 2:
                gridCols[1].set(qn('w:w'), str(col2_twips))

    def _replace_cover_page(self):
        """変換元の表紙をテンプレートの表紙で置き換える"""
        template_body = self.template_doc._body._body
        target_body = self.target_doc._body._body

        # テンプレートの表紙要素を取得
        template_cover = self._get_cover_page_elements(self.template_doc)

        # 変換元の表紙要素を取得
        target_cover = self._get_cover_page_elements(self.target_doc)

        # 変換元の表紙要素の最初の要素の前の位置を記録
        if target_cover:
            insert_position = list(target_body).index(target_cover[0])
        else:
            insert_position = 0

        # 変換元の表紙要素を削除
        for elem in target_cover:
            target_body.remove(elem)

        # テンプレートの表紙要素をコピーして挿入
        first_table = True
        for i, elem in enumerate(template_cover):
            new_elem = deepcopy(elem)
            # 変数を置換
            self._replace_variables_in_element(new_elem)

            # 最初のテーブルの1列目を45mm、2列目を135mmに設定
            if new_elem.tag == qn('w:tbl') and first_table:
                self._set_table_column_widths(new_elem, 45, 135)
                first_table = False

            target_body.insert(insert_position + i, new_elem)


def main():
    parser = argparse.ArgumentParser(
        description='DOCXファイルのヘッダー、フッター、表紙をテンプレートに従って変更します。',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用例:
    python top-page.py "sample/21GXP-D-001 ソフトウェア要求仕様書.docx" --DocCode D-001 --DocName ソフトウェア要求仕様書 --Version 1.0
        '''
    )

    parser.add_argument('source', type=str, help='変換元のDOCXファイルパス')
    parser.add_argument('--template', type=str, default='template.docx',
                        help='テンプレートファイルパス (default: template.docx)')
    parser.add_argument('--DocCode', type=str, default='D-000',
                        help='文書コード (default: D-000)')
    parser.add_argument('--DocName', type=str, default='仕様書',
                        help='文書名 (default: 仕様書)')
    parser.add_argument('--Version', type=str, default='1.0',
                        help='バージョン (default: 1.0)')

    args = parser.parse_args()

    # パスの処理
    source_path = Path(args.source)
    template_path = Path(args.template)

    # 変換元ファイルの存在確認
    if not source_path.exists():
        print(f'エラー: 変換元ファイルが見つかりません: {source_path}', file=sys.stderr)
        sys.exit(1)

    # テンプレートファイルの存在確認
    if not template_path.exists():
        print(f'エラー: テンプレートファイルが見つかりません: {template_path}', file=sys.stderr)
        sys.exit(1)

    # 出力パスの設定（sourceの親ディレクトリ/changed/ファイル名）
    output_path = source_path.parent / 'changed' / source_path.name

    # 変数の設定
    variables = {
        'DocCode': args.DocCode,
        'DocName': args.DocName,
        'Version': args.Version,
    }

    # フォーマット実行
    formatter = DocxFormatter(template_path, variables)
    formatter.format(source_path, output_path)


if __name__ == '__main__':
    main()
