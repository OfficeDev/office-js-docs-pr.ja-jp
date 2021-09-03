---
title: JavaScript API を使用してコメントExcelする
description: API を使用してコメントとコメント スレッドを追加、削除、および編集する方法について説明します。
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 16569bc1d72391dff0ac35a48e45470ff90852f8
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868653"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>JavaScript API を使用してコメントExcelする

この記事では、JavaScript API を使用してブック内のコメントを追加、読み取り、変更、削除するExcel説明します。 コメント機能の詳細については、「コメントとメモを挿入する」の記事[Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)できます。

JavaScript API Excel、コメントには、1 つの初期コメントと接続されたスレッドディスカッションの両方が含まれます。 これは、個々のセルに関連付けされます。 十分なアクセス許可を持つブックを表示しているユーザーは、コメントに返信できます。 [Comment オブジェクト](/javascript/api/excel/excel.comment)は、これらの返信を[CommentReply オブジェクトとして格納](/javascript/api/excel/excel.commentreply)します。 コメントはスレッドであり、スレッドには開始点として特別なエントリが必要と考える必要があります。

![コメントExcel"Comment" というラベルが付き、"Comment.replies[0]" と "Comment.replies[1]" というラベルが付けされています。](../images/excel-comments.png)

ブック内のコメントは、プロパティによって追跡 `Workbook.comments` されます。 これには、ユーザーによって作成されたコメントだけでなく、アドインによって作成されたコメントも含まれます。 `Workbook.comments` プロパティは、[Comment](/javascript/api/excel/excel.comment) オブジェクトのコレクションを含む [CommentCollection](/javascript/api/excel/excel.commentcollection) オブジェクトです。 コメントはワークシート レベルでも [アクセス](/javascript/api/excel/excel.worksheet) できます。 この記事のサンプルは、ブック レベルのコメントを扱いますが、プロパティを使用するために簡単に変更 `Worksheet.comments` できます。

## <a name="add-comments"></a>コメントを追加する

ブックに `CommentCollection.add` コメントを追加するには、このメソッドを使用します。 このメソッドは、最大 3 つのパラメーターを受け取ります。

- `cellAddress`: コメントが追加されるセル。 文字列または Range オブジェクトを [指定](/javascript/api/excel/excel.range) できます。 範囲は 1 つのセルである必要があります。
- `content`: コメントのコンテンツ。 テキスト形式のコメントには文字列を使用します。 メンション [付きコメントには CommentRichContent](/javascript/api/excel/excel.commentrichcontent) オブジェクト [を使用します](#mentions)。
- `contentType`: コンテンツ [の種類を](/javascript/api/excel/excel.contenttype) 指定する ContentType 列挙。 既定値は `ContentType.plain` です。

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
> アドインによって追加されたコメントは、そのアドインの現在のユーザーに属性付けされます。

### <a name="add-comment-replies"></a>コメント返信の追加

オブジェクト `Comment` は、ゼロ以上の返信を含むコメント スレッドです。 `Comment` オブジェクトには `replies` プロパティがあり、これは [CommentReply](/javascript/api/excel/excel.commentreply) オブジェクトを含む [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) です。 コメントに返信を追加するには、`CommentReplyCollection.add` メソッドを使用して、返信のテキストを渡します。 返信は、追加された順に表示されます。 これらは、アドインの現在のユーザーにも属性付けされます。

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

### <a name="edit-comment-replies"></a>コメント返信の編集

コメント返信を編集するには、そのプロパティを設定 `CommentReply.content` します。

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

コメントを削除するには、メソッドを `Comment.delete` 使用します。 コメントを削除すると、そのコメントに関連付けられた返信も削除されます。

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>コメントの返信を削除する

コメントの返信を削除するには、メソッドを使用 `CommentReply.delete` します。

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a>コメント スレッドの解決

コメント スレッドには、解決されたかどうかを示す構成可能なブール `resolved` 値があります。 値は `true` 、コメント スレッドが解決された値を意味します。 値は、 `false` コメント スレッドが新規または再オープンされた値を意味します。

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

コメントの返信には readonly プロパティ `resolved` があります。 その値は、スレッドの残りの値と常に等しくなります。

## <a name="comment-metadata"></a>コメント メタデータ

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

### <a name="comment-reply-metadata"></a>コメント返信メタデータ

コメント返信には、最初のコメントと同じ種類のメタデータが格納されます。

次のサンプルは、作成者の電子メール、作成者の名前、および A2 での最新のコメント返信の作成日を表示する方法を **示しています**。

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

## <a name="mentions"></a>メンション

[メンションは](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) 、コメント内の同僚にタグを付けするために使用されます。 これにより、コメントのコンテンツと一緒に通知が送信されます。 アドインは、ユーザーに代わってこれらのメンションを作成できます。

CommentRichContent オブジェクトを使用して、メンションを含むコメント [を作成する必要](/javascript/api/excel/excel.commentrichcontent) があります。 1 `CommentCollection.add` つ `CommentRichContent` 以上のメンションを含む呼び出しを実行し、パラメーター `ContentType.mention` として指定 `contentType` します。 また `content` 、テキストにメンションを挿入するには、文字列を書式設定する必要があります。 メンションの形式は次の形式です `<at id="{replyIndex}">{mentionName}</at>` 。

> [!NOTE]
> 現在、メンションリンクのテキストとして使用できるのは、メンションの正確な名前のみです。 名前の短縮バージョンのサポートは、後で追加されます。

次の例は、1 つのメンションを含むコメントを示しています。

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

## <a name="comment-events"></a>コメント イベント

アドインは、コメントの追加、変更、削除をリッスンできます。 [コメント イベントは](/javascript/api/excel/excel.commentcollection#event-details) 、オブジェクトで発生 `CommentCollection` します。 コメント イベントをリッスンするには、、 `onAdded` `onChanged` 、、またはコメント イベント `onDeleted` ハンドラーを登録します。 コメント イベントが検出された場合は、このイベント ハンドラーを使用して、追加、変更、または削除されたコメントに関するデータを取得します。 イベント `onChanged` は、コメントの返信の追加、変更、および削除も処理します。 

各コメント イベントは、複数の追加、変更、または削除が同時に実行されると 1 回だけトリガーされます。 すべての [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)オブジェクト [、CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)オブジェクト、 [および CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) オブジェクトには、イベント アクションをコメント コレクションにマップするコメント ID の配列が含まれています。

イベント ハンドラーの[Excel登録、イベントの処理、およびイベント ハンドラーの](excel-add-ins-events.md)削除の詳細については、「JavaScript API を使用したイベントの処理」の記事を参照してください。 

### <a name="comment-addition-events"></a>コメントの追加イベント 
イベント `onAdded` は、1 つ以上の新しいコメントがコメント コレクションに追加されるとトリガーされます。 このイベントは *、返信* がコメント スレッドに追加された場合にはトリガーされません ([](#comment-change-events)コメントの返信イベントについては、「コメント変更イベント」を参照してください)。

次のサンプルは、イベント ハンドラーを登録し、オブジェクトを使用して追加されたコメントの配列 `onAdded` `CommentAddedEventArgs` `commentDetails` を取得する方法を示しています。

> [!NOTE]
> このサンプルは、1 つのコメントが追加された場合にのみ機能します。 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    return context.sync();
});

function commentAdded() {
    Excel.run(function (context) {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        var addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the added comment's data.
            console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
            return context.sync();
        });            
    });
}
```

### <a name="comment-change-events"></a>コメント変更イベント 
コメント `onChanged` イベントは、次のシナリオでトリガーされます。

- コメントのコンテンツが更新されます。
- コメント スレッドが解決されます。
- コメント スレッドが再び開きます。
- コメント スレッドに返信が追加されます。
- コメント スレッドで返信が更新されます。
- コメント スレッドで返信が削除されます。

次のサンプルは、イベント ハンドラーを登録し、オブジェクトを使用して変更されたコメントの `onChanged` `CommentChangedEventArgs` `commentDetails` 配列を取得する方法を示しています。

> [!NOTE]
> このサンプルは、1 つのコメントが変更された場合にのみ機能します。 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    return context.sync();
});    

function commentChanged() {
    Excel.run(function (context) {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        var changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the changed comment's data.
            console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}`. Updated comment content: ${changedComment.content}`. Comment author: ${changedComment.authorName}`);
            return context.sync();
        });
    });
}
```

### <a name="comment-deletion-events"></a>コメント削除イベント
コメント `onDeleted` がコメント コレクションから削除されると、イベントがトリガーされます。 コメントが削除された後、そのメタデータは使用できなくなりました。 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)オブジェクトは、アドインが個々のコメントを管理している場合に備え、コメントの ID を提供します。

次のサンプルは、イベント ハンドラーを登録し、オブジェクトを使用して削除されたコメントの配列 `onDeleted` `CommentDeletedEventArgs` `commentDetails` を取得する方法を示しています。

> [!NOTE]
> このサンプルは、1 つのコメントが削除された場合にのみ機能します。 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    return context.sync();
});

function commentDeleted() {
    Excel.run(function (context) {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してブックを操作する](excel-add-ins-workbooks.md)
- [Excel JavaScript API を使用してイベントを操作する](excel-add-ins-events.md)
- [コメントとメモをページに挿入Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
