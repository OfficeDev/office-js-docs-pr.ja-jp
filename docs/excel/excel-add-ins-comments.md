---
title: Excel JavaScript API を使用してコメントを操作する
description: Api を使用してコメントおよびコメントスレッドを追加、削除、および編集する方法について説明します。
ms.date: 03/17/2020
localization_priority: Normal
ms.openlocfilehash: 275828915730d3438101315ee28bf76aa8b8bf3f
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890571"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してコメントを操作する

この記事では、Excel JavaScript API を使用してブック内のコメントを追加、読み取り、変更、および削除する方法について説明します。 コメント機能の詳細については、「 [Excel 記事のコメントとメモを挿入する」](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)を参照してください。

Excel JavaScript API では、コメントには単一の最初のコメントと接続されたスレッドのディスカッションの両方が含まれます。 個別のセルに関連付けられています。 十分な権限があるブックを表示するユーザーは、コメントに返信できます。 Comment オブジェクトは、これらの返信を[コメント](/javascript/api/excel/excel.comment)[返信](/javascript/api/excel/excel.commentreply)オブジェクトとして格納します。 コメントはスレッドと考えてください。スレッドには、開始点として特別なエントリが必要です。

![「Comment」というラベルが付けられた、"comment" というラベルが付いた Excel コメント。「comment [0]」と「Comment [1]」。](../images/excel-comments.png)

ブック内のコメントは`Workbook.comments`プロパティによって追跡されます。 これには、ユーザーによって作成されたコメントだけでなく、アドインによって作成されたコメントも含まれます。 `Workbook.comments` プロパティは、[Comment](/javascript/api/excel/excel.comment) オブジェクトのコレクションを含む [CommentCollection](/javascript/api/excel/excel.commentcollection) オブジェクトです。 コメントには、[ワークシート](/javascript/api/excel/excel.worksheet)レベルでアクセスすることもできます。 この記事のサンプルでは、ブックレベルでコメントを使用していますが、 `Worksheet.comments`プロパティを使用するために簡単に変更することができます。

## <a name="add-comments"></a>コメントを追加する

メソッドを`CommentCollection.add`使用して、ブックにコメントを追加します。 このメソッドは、次の3つのパラメーターを取ります。

- `cellAddress`: コメントが追加されるセルを指定します。 文字列または[Range](/javascript/api/excel/excel.range)オブジェクトのいずれかを指定できます。 範囲は1つのセルである必要があります。
- `content`: コメントの内容。 テキスト形式のコメントには文字列を使用します。 [メンション](#mentions-online-only)付きのコメントには、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用します。 
- `contentType`: コンテンツの種類を指定する[ContentType](/javascript/api/excel/excel.contenttype)列挙。 既定値は `ContentType.plain` です。

次のコード例は、コメントをセル **A2** に追加します。

```js
Excel.run(function (context) {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    return context.sync();
});
```

> [!NOTE]
> アドインによって追加されたコメントは、そのアドインの現在のユーザーによって作成されます。

### <a name="add-comment-replies"></a>コメントの返信を追加する

`Comment`オブジェクトは、0個以上の返信を含むコメントスレッドです。 `Comment` オブジェクトには `replies` プロパティがあり、これは [CommentReply](/javascript/api/excel/excel.commentreply) オブジェクトを含む [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) です。 コメントに返信を追加するには、`CommentReplyCollection.add` メソッドを使用して、返信のテキストを渡します。 返信は、追加された順に表示されます。 また、アドインの現在のユーザーにも属性があります。

次のコード サンプルは、ブックの最初のコメントに返信を追加します。

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a>コメントの編集

コメントまたはコメントの返信を編集するには、その `Comment.content` プロパティまたは `CommentReply.content` プロパティを設定します。

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a>コメントの返信を編集する

コメントの返信を編集するには`CommentReply.content` 、そのプロパティを設定します。

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a>コメントの削除

コメントを削除するには`Comment.delete` 、メソッドを使用します。 コメントを削除すると、そのコメントに関連付けられている返信も削除されます。

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>コメントの返信を削除する

コメントの返信を削除するには`CommentReply.delete` 、メソッドを使用します。

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads-preview"></a>コメントスレッドを解決する ([プレビュー](../reference/requirement-sets/excel-preview-apis.md)) 

コメントスレッドには、解決可能かどう`resolved`かを示す、構成可能なブール値があります。 の`true`値は、コメントスレッドが解決されたことを意味します。 の`false`値は、コメントスレッドが新規または再オープンのいずれかであることを意味します。

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

コメントの返信には`resolved` 、readonly プロパティがあります。 この値は、常にスレッドの残りの部分と同じです。

## <a name="comment-metadata"></a>コメントのメタデータ

各コメントには、作成者や作成日などの作成に関するメタデータが含まれています。 アドインによって作成されたコメントは、現在のユーザーによって作成されたものと見なされます。

次のサンプルは、**A2** に作成者のメール、作成者の名前、コメントの作成日を表示する方法を示しています。

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

### <a name="comment-reply-metadata"></a>コメントの返信メタデータ

コメントの返信は、最初のコメントと同じ種類のメタデータを格納します。

次の例は、作成者の電子メール、作成者の名前、および**A2**における最新のコメントの返信の作成日を表示する方法を示しています。

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    var replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    return context.sync().then(function () {
        // Get the last comment reply in the comment thread.
        var reply = comment.replies.getItemAt(replyCount.value - 1);
        reply.load(["authorEmail", "authorName", "creationDate"]);
        // Sync to load the reply metadata to print.
        return context.sync().then(function () {
            console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
            return context.sync();
        });
    });
});
```

## <a name="mentions-online-only"></a>メンション ([オンラインのみ](../reference/requirement-sets/excel-api-online-requirement-set.md)) 

> [!NOTE]
> コメントコメント Api は、現在、パブリックプレビューでのみ利用可能です。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> コメントメンションは、現在 web 上の Excel でのみサポートされています。

[メンション](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd)は、コメント内の仕事仲間にタグ付けするために使用されます。 これにより、それらの通知がコメントの内容と共に送信されます。 アドインは、ユーザーの代わりにこれらのメンションを作成できます。

[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用して、メンションを含むコメントを作成する必要があります。 1 `CommentCollection.add`つ以上`CommentRichContent`のメンションを含むを呼び出し`ContentType.mention` 、 `contentType`パラメーターとしてを指定します。 `content`文字列をテキストに挿入するには、文字列を書式設定する必要もあります。 メンションの形式は、 `<at id="{replyIndex}">{mentionName}</at>`です。

> こと現時点では、メンションリンクのテキストとして、メンションの正確な名前のみを使用できます。 名前の短縮バージョンのサポートは、後で追加されます。

次の例は、1つのメンション付きのコメントを示しています。

```js
Excel.run(function (context) {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    var mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    var commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    return context.sync();
});
```

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してブックを操作する](excel-add-ins-workbooks.md)
- [Excel でコメントやメモを挿入する](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
