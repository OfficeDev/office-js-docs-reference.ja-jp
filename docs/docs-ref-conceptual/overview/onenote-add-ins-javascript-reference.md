# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="3641b-101">OneNote の JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="3641b-101">OneNote JavaScript API overview</span></span>

<span data-ttu-id="3641b-102">適用対象:OneNote Online</span><span class="sxs-lookup"><span data-stu-id="3641b-102">Applies to: OneNote Online</span></span>

<span data-ttu-id="3641b-103">以下のリンクは、API で使用できる高レベルの OneNote オブジェクトを示しています。</span><span class="sxs-lookup"><span data-stu-id="3641b-103">The following links show the high level OneNote objects available in the API.</span></span> <span data-ttu-id="3641b-104">各オブジェクトのページのリンクには、プロパティ、イベント、およびオブジェクトの使用可能なメソッドの説明が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3641b-104">Each object page link contains a description of the properties, events, and methods available on the object.</span></span> <span data-ttu-id="3641b-105">リンクを調べて詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="3641b-105">Explore these links to learn more.</span></span> 
    
- <span data-ttu-id="3641b-106">[Application](/javascript/api/onenote/onenote.application):グローバルにアドレス可能な OneNote オブジェクト (アクティブなノートブック、アクティブなセクションなど) すべてへのアクセスに使用する最上位のオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3641b-106">[Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.</span></span>

- <span data-ttu-id="3641b-p102">[Notebook](/javascript/api/onenote/onenote.notebook):ノートブックです。ノートブックには、セクション グループとセクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3641b-p102">[Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.</span></span>
    - <span data-ttu-id="3641b-109">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection):ノートブックのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3641b-109">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.</span></span>

- <span data-ttu-id="3641b-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup):セクション グループです。セクション グループには、セクション グループとセクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3641b-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.</span></span>
    - <span data-ttu-id="3641b-112">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection):セクション グループのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3641b-112">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.</span></span>

- <span data-ttu-id="3641b-p104">[Section](/javascript/api/onenote/onenote.section):セクションです。セクションには、ページが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3641b-p104">[Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.</span></span>
    - <span data-ttu-id="3641b-115">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection):セクションのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3641b-115">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.</span></span>

- <span data-ttu-id="3641b-p105">[Page](/javascript/api/onenote/onenote.page):ページです。ページには、PageContent オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3641b-p105">[Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.</span></span>
    - <span data-ttu-id="3641b-118">[PageCollection](/javascript/api/onenote/onenote.pagecollection):ページのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3641b-118">[PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.</span></span>

- <span data-ttu-id="3641b-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent):Outline や Image などのコンテンツの種類を含むページの最上位の領域です。PageContent オブジェクトは、ページ上の位置を指定できます。</span><span class="sxs-lookup"><span data-stu-id="3641b-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.</span></span>
    - <span data-ttu-id="3641b-121">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection):PageContent オブジェクトのコレクションで、ページのコンテンツを表します。</span><span class="sxs-lookup"><span data-stu-id="3641b-121">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.</span></span>

- <span data-ttu-id="3641b-p107">[Outline](/javascript/api/onenote/onenote.outline):Paragraph オブジェクトのコンテナーです。Outline は、PageContent オブジェクトの直接の子です。</span><span class="sxs-lookup"><span data-stu-id="3641b-p107">[Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.</span></span>

- <span data-ttu-id="3641b-p108">[Image](/javascript/api/onenote/onenote.image):Image オブジェクトです。Image は、PageContent オブジェクトまたは Paragraph の直接の子にすることができます。</span><span class="sxs-lookup"><span data-stu-id="3641b-p108">[Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.</span></span>

- <span data-ttu-id="3641b-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph):ページに表示されるコンテンツのコンテナーです。Paragraph は、Outline の直接の子です。</span><span class="sxs-lookup"><span data-stu-id="3641b-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.</span></span>
    - <span data-ttu-id="3641b-128">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection):Outline 内の Paragraph オブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3641b-128">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.</span></span>

- <span data-ttu-id="3641b-129">[RichText](/javascript/api/onenote/onenote.richtext):RichText オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3641b-129">[RichText](/javascript/api/onenote/onenote.richtext): A RichText object.</span></span>

- <span data-ttu-id="3641b-130">[Table](/javascript/api/onenote/onenote.table):TableRow オブジェクトのコンテナーです。</span><span class="sxs-lookup"><span data-stu-id="3641b-130">[Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.</span></span>

- <span data-ttu-id="3641b-131">[TableRow](/javascript/api/onenote/onenote.tablerow):TableCell オブジェクトのコンテナーです。</span><span class="sxs-lookup"><span data-stu-id="3641b-131">[TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.</span></span>
    - <span data-ttu-id="3641b-132">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection):Table 内の TableRow オブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3641b-132">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.</span></span>
 
- <span data-ttu-id="3641b-133">[TableCell](/javascript/api/onenote/onenote.tablecell):Paragraph オブジェクトのコンテナーです。</span><span class="sxs-lookup"><span data-stu-id="3641b-133">[TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.</span></span>
    - <span data-ttu-id="3641b-134">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection):TableRow 内の TableCell オブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3641b-134">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.</span></span>

## <a name="onenote-javascript-api-reference"></a><span data-ttu-id="3641b-135">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="3641b-135">OneNote JavaScript API reference</span></span>

<span data-ttu-id="3641b-136">OneNote の JavaScript API の詳細については、 [OneNote の JavaScript API リファレンス ドキュメント](/javascript/api/onenote)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3641b-136">For detailed information about OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="3641b-137">関連項目</span><span class="sxs-lookup"><span data-stu-id="3641b-137">See also</span></span>

- [<span data-ttu-id="3641b-138">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="3641b-138">OneNote JavaScript API programming overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [<span data-ttu-id="3641b-139">最初の OneNote 用アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="3641b-139">Build your first OneNote add-in</span></span>](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [<span data-ttu-id="3641b-140">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="3641b-140">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="3641b-141">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="3641b-141">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
