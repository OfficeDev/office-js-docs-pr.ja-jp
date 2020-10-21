---
title: Excel JavaScript API を使用してコメントを操作する
description: Api を使用してコメントおよびコメントスレッドを追加、削除、および編集する方法について説明します。
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 00f7dd22fb2148902152197521098482071e5284
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626422"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してコメントを操作する

この記事では、Excel JavaScript API を使用してブック内のコメントを追加、読み取り、変更、および削除する方法について説明します。 コメント機能の詳細については、「 [Excel 記事のコメントとメモを挿入する」](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) を参照してください。

Excel JavaScript API では、コメントには単一の最初のコメントと接続されたスレッドのディスカッションの両方が含まれます。 個別のセルに関連付けられています。 十分な権限があるブックを表示するユーザーは、コメントに返信できます。 Comment オブジェクトは、これらの返信を[コメント](/javascript/api/excel/excel.comment)[返信](/javascript/api/excel/excel.commentreply)オブジェクトとして格納します。 コメントはスレッドと考えてください。スレッドには、開始点として特別なエントリが必要です。

![「Comment」というラベルが付けられた、"comment" というラベルが付いた Excel コメント。「comment [0]」と「Comment [1]」。](../images/excel-comments.png)

ブック内のコメントはプロパティによって追跡され `Workbook.comments` ます。 これには、ユーザーによって作成されたコメントだけでなく、アドインによって作成されたコメントも含まれます。 `Workbook.comments` プロパティは、[Comment](/javascript/api/excel/excel.comment) オブジェクトのコレクションを含む [CommentCollection](/javascript/api/excel/excel.commentcollection) オブジェクトです。 コメントには、 [ワークシート](/javascript/api/excel/excel.worksheet) レベルでアクセスすることもできます。 この記事のサンプルでは、ブックレベルでコメントを使用していますが、プロパティを使用するために簡単に変更することができ `Worksheet.comments` ます。

## <a name="add-comments"></a>コメントを追加する

メソッドを使用して、 `CommentCollection.add` ブックにコメントを追加します。 このメソッドは、次の3つのパラメーターを取ります。

- `cellAddress`: コメントが追加されるセルを指定します。 文字列または [Range](/javascript/api/excel/excel.range) オブジェクトのいずれかを指定できます。 範囲は1つのセルである必要があります。
- `content`: コメントの内容。 テキスト形式のコメントには文字列を使用します。 [メンション](#mentions)付きのコメントには、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用します。
- `contentType`: コンテンツの種類を指定する [ContentType](/javascript/api/excel/excel.contenttype) 列挙。 既定値は `ContentType.plain` です。

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

コメントの返信を編集するには、そのプロパティを設定 `CommentReply.content` します。

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

コメントを削除するには、メソッドを使用し `Comment.delete` ます。 コメントを削除すると、そのコメントに関連付けられている返信も削除されます。

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>コメントの返信を削除する

コメントの返信を削除するには、メソッドを使用し `CommentReply.delete` ます。

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a>コメントスレッドを解決する

コメントスレッドには、解決可能かどうかを示す、構成可能なブール値があり `resolved` ます。 の値は、 `true` コメントスレッドが解決されたことを意味します。 の値は、 `false` コメントスレッドが新規または再オープンのいずれかであることを意味します。

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

コメントの返信には、readonly プロパティがあり `resolved` ます。 この値は、常にスレッドの残りの部分と同じです。

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

次の例は、作成者の電子メール、作成者の名前、および **A2**における最新のコメントの返信の作成日を表示する方法を示しています。

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

[メンション](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) は、コメント内の仕事仲間にタグ付けするために使用されます。 これにより、それらの通知がコメントの内容と共に送信されます。 アドインは、ユーザーの代わりにこれらのメンションを作成できます。

[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)オブジェクトを使用して、メンションを含むコメントを作成する必要があります。 1つ以上のメンションを含むを呼び出し、 `CommentCollection.add` `CommentRichContent` `ContentType.mention` パラメーターとしてを指定し `contentType` ます。 `content`文字列をテキストに挿入するには、文字列を書式設定する必要もあります。 メンションの形式は、 `<at id="{replyIndex}">{mentionName}</at>` です。

> [!NOTE]
> 現時点では、メンションリンクのテキストとして、メンションの正確な名前のみを使用できます。 名前の短縮バージョンのサポートは、後で追加されます。

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

## <a name="comment-events"></a>コメントイベント

アドインは、コメントの追加、変更、および削除を聞くことができます。 [Comment イベント](/javascript/api/excel/excel.commentcollection#event-details) は、オブジェクトに対して発生 `CommentCollection` します。 Comment イベントをリッスンするには、、、 `onAdded` `onChanged` またはの `onDeleted` コメントイベントハンドラーを登録します。 コメントイベントが検出されたときに、追加、変更、または削除されたコメントに関するデータを取得するには、このイベントハンドラーを使用します。 この `onChanged` イベントは、コメントの返信の追加、変更、および削除も処理します。 

各 comment イベントは、同時に複数の追加、変更、または削除が実行された場合にのみトリガーされます。 [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)、 [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)、および[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)のすべてのオブジェクトには、イベントアクションをコメントのコレクションにマップするためのコメント id の配列が含まれています。

イベントハンドラーの登録、イベントの処理、イベントハンドラーの削除に関する追加情報については、「 [Excel JAVASCRIPT API を使用してイベント](excel-add-ins-events.md) を処理する」の記事を参照してください。 

### <a name="comment-addition-events"></a>コメントの追加イベント 
この `onAdded` イベントは、コメントのコレクションに1つまたは複数の新しいコメントが追加されると発生します。 このイベントは、コメントスレッドに返信が追加されたときには発生し *ません* (コメントの返信イベントについては、「 [コメント変更イベント](#comment-change-events) 」を参照してください)。

次の例は、イベントハンドラーを登録し、そのオブジェクトを使用して追加されたコメントの配列を取得する方法を示して `onAdded` `CommentAddedEventArgs` `commentDetails` います。

> [!NOTE]
> このサンプルは、1つのコメントが追加された場合にのみ機能します。 

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
`onChanged`Comment イベントは、次のシナリオでトリガーされます。

- コメントの内容が更新されます。
- コメントスレッドが解決されます。
- コメントスレッドが再度開かれています。
- コメントスレッドに返信が追加されます。
- コメントスレッド内の返信が更新されます。
- コメントスレッド内の返信が削除されます。

次の例は、イベントハンドラーを登録し、そのオブジェクトを使用して、変更されたコメントの配列を取得する方法を示して `onChanged` `CommentChangedEventArgs` `commentDetails` います。

> [!NOTE]
> このサンプルは、1つのコメントが変更された場合にのみ機能します。 

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
コメントの `onDeleted` コレクションからコメントが削除されると、イベントがトリガーされます。 コメントが削除されると、そのメタデータは使用できなくなります。 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)オブジェクトは、アドインが個々のコメントを管理している場合に、コメント id を提供します。

次の例は、イベントハンドラーを登録し、そのオブジェクトを使用して、削除されたコメントの配列を取得する方法を示して `onDeleted` `CommentDeletedEventArgs` `commentDetails` います。

> [!NOTE]
> このサンプルは、1つのコメントが削除された場合にのみ機能します。 

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

- [Office アドインでの Excel JavaScript オブジェクトモデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してブックを操作する](excel-add-ins-workbooks.md)
- [Excel JavaScript API を使用してイベントを操作する](excel-add-ins-events.md)
- [Excel でコメントやメモを挿入する](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
