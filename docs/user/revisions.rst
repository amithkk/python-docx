.. _revisions:

Working with Tracked Changes (Revisions)
========================================

Word allows *track changes* (also known as *revisions*) to be enabled on a document.
This feature records insertions and deletions made to the document, showing who made
each change and when. This is commonly used for collaborative editing and review
workflows.

When track changes is enabled:

- Inserted text is marked with the ``<w:ins>`` element
- Deleted text is marked with the ``<w:del>`` element
- Each revision records the author, date, and a unique revision ID

.. note::

    *python-docx* supports creating and reading tracked changes, as well as accepting
    or rejecting individual revisions programmatically.


Enabling Track Changes
----------------------

Track changes mode is controlled via the document settings::

    >>> from docx import Document
    >>> document = Document()
    >>> document.settings.track_revisions = True
    >>> document.settings.track_revisions
    True

When ``track_revisions`` is ``True``, Word will track any subsequent changes made in
the Word application. Changes made programmatically via *python-docx* must be
explicitly marked as tracked using the methods described below.


Find and Replace with Track Changes
-----------------------------------

The most common use case is performing a find-and-replace operation where the changes
are tracked. The :meth:`.Document.find_and_replace_tracked` method handles this::

    >>> document = Document("contract.docx")
    >>> document.settings.track_revisions = True
    >>> count = document.find_and_replace_tracked(
    ...     search_text="Acme Corp",
    ...     replace_text="NewCo Inc",
    ...     author="Legal Team",
    ...     comment="Company name updated per merger agreement",
    ... )
    >>> print(f"Replaced {count} occurrences")
    Replaced 15 occurrences
    >>> document.save("contract_revised.docx")

This method:

- Searches all paragraphs in the document body and tables
- Replaces only the specific text (word-level), preserving surrounding formatting
- Creates tracked deletions for the old text and tracked insertions for the new text
- Optionally attaches a comment to each replacement explaining the change

For more control, you can use :meth:`.Paragraph.replace_tracked` on individual
paragraphs::

    >>> for paragraph in document.paragraphs:
    ...     if "confidential" in paragraph.text.lower():
    ...         paragraph.replace_tracked("draft", "final", author="Editor")


Adding Tracked Insertions
-------------------------

To add new text as a tracked insertion::

    >>> paragraph = document.add_paragraph("This is existing text. ")
    >>> tracked = paragraph.add_run_tracked(
    ...     text="This was added later.",
    ...     author="John Smith",
    ... )
    >>> tracked
    <docx.revision.TrackedInsertion object at 0x...>
    >>> tracked.author
    'John Smith'
    >>> tracked.text
    'This was added later.'

The ``add_run_tracked`` method wraps the new text in a ``<w:ins>`` element, marking
it as inserted content that will appear in Word's track changes view.


Creating Tracked Deletions
--------------------------

To mark existing text as deleted (without actually removing it)::

    >>> paragraph = document.add_paragraph("Delete this text please.")
    >>> run = paragraph.runs[0]
    >>> tracked = run.delete_tracked(author="Editor")
    >>> tracked
    <docx.revision.TrackedDeletion object at 0x...>
    >>> tracked.text
    'Delete this text please.'

The text remains in the document but is wrapped in a ``<w:del>`` element. In Word,
this text appears with strikethrough formatting.


Iterating Over Revisions
------------------------

To access tracked changes in a paragraph, use ``iter_inner_content`` with
``include_revisions=True``::

    >>> from docx.revision import TrackedInsertion, TrackedDeletion
    >>> for item in paragraph.iter_inner_content(include_revisions=True):
    ...     if isinstance(item, TrackedInsertion):
    ...         print(f"INSERTED by {item.author}: {item.text}")
    ...     elif isinstance(item, TrackedDeletion):
    ...         print(f"DELETED by {item.author}: {item.text}")
    ...     else:
    ...         print(f"Normal text: {item.text}")


Accepting and Rejecting Changes
-------------------------------

Individual revisions can be accepted or rejected programmatically::

    >>> # Accept an insertion (keeps the inserted text)
    >>> tracked_insertion.accept()

    >>> # Reject an insertion (removes the inserted text)
    >>> tracked_insertion.reject()

    >>> # Accept a deletion (removes the deleted text)
    >>> tracked_deletion.accept()

    >>> # Reject a deletion (restores the deleted text)
    >>> tracked_deletion.reject()


TrackedInsertion and TrackedDeletion Properties
-----------------------------------------------

Both ``TrackedInsertion`` and ``TrackedDeletion`` objects provide these properties:

``author``
    The name of the author who made the change (read/write).

``date``
    The date and time of the change as a ``datetime`` object (read-only).

``revision_id``
    The unique identifier for this revision (read/write).

``text``
    The text content of the revision (read-only).

``runs``
    A list of ``Run`` objects contained in the revision.

``is_run_level``
    ``True`` if the revision contains runs (inline content).

``is_block_level``
    ``True`` if the revision contains paragraphs or tables (block content).


Example: Processing a Document with Track Changes
-------------------------------------------------

Here's a complete example that processes an existing document with track changes::

    >>> from docx import Document
    >>> from docx.revision import TrackedInsertion, TrackedDeletion

    >>> document = Document("reviewed_document.docx")

    >>> # Count revisions
    >>> insertions = 0
    >>> deletions = 0

    >>> for paragraph in document.paragraphs:
    ...     for item in paragraph.iter_inner_content(include_revisions=True):
    ...         if isinstance(item, TrackedInsertion):
    ...             insertions += 1
    ...             print(f"[+] {item.author}: {item.text[:50]}...")
    ...         elif isinstance(item, TrackedDeletion):
    ...             deletions += 1
    ...             print(f"[-] {item.author}: {item.text[:50]}...")

    >>> print(f"\nTotal: {insertions} insertions, {deletions} deletions")


Example: Bulk Accept All Changes
--------------------------------

To accept all tracked changes in a document::

    >>> from docx import Document
    >>> from docx.revision import TrackedInsertion, TrackedDeletion

    >>> document = Document("reviewed_document.docx")

    >>> for paragraph in document.paragraphs:
    ...     for item in list(paragraph.iter_inner_content(include_revisions=True)):
    ...         if isinstance(item, (TrackedInsertion, TrackedDeletion)):
    ...             item.accept()

    >>> document.save("accepted_document.docx")

Note the use of ``list()`` to materialize the iterator before modifying the document,
as accepting/rejecting changes modifies the underlying XML.