# <a name="onenote-javascript-api-overview"></a>OneNote の JavaScript API の概要

適用対象: OneNote Online

以下のリンクは、API で使用できるハイレベルの OneNote オブジェクトを示しています。 オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、イベント、メソッドの説明が含まれています。 リンクを調べて詳細を確認してください。 
    
- [Application](/javascript/api/onenote/onenote.application): グローバルにアドレス可能なすべての OneNote オブジェクト (アクティブなノートブック、アクティブなセクションなど) へのアクセスに使われる最上位のオブジェクトです。

- [Notebook](/javascript/api/onenote/onenote.notebook): ノートブックです。ノートブックには、セクション グループとセクションが含まれます。
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): ノートブックのコレクションです。

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup): セクション グループです。セクション グループには、セクション グループとセクションが含まれます。
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): セクション グループのコレクションです。

- [Section](/javascript/api/onenote/onenote.section): セクションです。セクションには、ページが含まれます。
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection): セクションのコレクションです。

- [Page](/javascript/api/onenote/onenote.page): ページです。ページには、PageContent オブジェクトが含まれます。
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection): ページのコレクションです。

- [PageContent](/javascript/api/onenote/onenote.pagecontent): Outline や Image などのコンテンツの種類を含むページの最上位の領域です。PageContent オブジェクトは、ページ上の位置を指定できます。
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): PageContent オブジェクトのコレクションで、ページのコンテンツを表します。

- [Outline](/javascript/api/onenote/onenote.outline): Paragraph オブジェクトのコンテナーです。Outline は、PageContent オブジェクトの直接の子です。

- [Image](/javascript/api/onenote/onenote.image): Image オブジェクトです。Image は、PageContent オブジェクトまたは Paragraph の直接の子にすることができます。

- [Paragraph](/javascript/api/onenote/onenote.paragraph): ページに表示されるコンテンツのコンテナーです。Paragraph は、Outline の直接の子です。
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): Outline 内の Paragraph オブジェクトのコレクションです。

- [RichText](/javascript/api/onenote/onenote.richtext): RichText オブジェクトです。

- [Table](/javascript/api/onenote/onenote.table): TableRow オブジェクトのコンテナーです。

- [TableRow](/javascript/api/onenote/onenote.tablerow): TableCell オブジェクトのコンテナーです。
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): Table 内の TableRow オブジェクトのコレクションです。
 
- [TableCell](/javascript/api/onenote/onenote.tablecell): Paragraph オブジェクトのコンテナーです。
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): TableRow 内の TableCell オブジェクトのコレクションです。

## <a name="onenote-javascript-api-reference"></a>OneNote JavaScript API リファレンス

OneNote の JavaScript API の詳細については、 [OneNote の JavaScript API 参照ドキュメント](/javascript/api/onenote)を参照してください。

## <a name="see-also"></a>関連項目

- [OneNote の JavaScript API のプログラミングの概要](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [最初の OneNote 用アドインをビルドする](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
