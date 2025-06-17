/**
* @OnlyCurrentDoc
*
* The above tag is used to prevent the script from accessing other documents
* that the user has not explicitly authorized.
*/

/**
* Creates a sheet name using the current date
* @return {string} Sheet name in format "Papers MMM DD YYYY"
*/
function createSheetName() {
    const now = new Date();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `Papers ${months[now.getMonth()]} ${now.getDate()} ${now.getFullYear()}`;
}

/**
* Gets the sheet for a specific week
* @param {Date} date - The date to get the week sheet for
* @return {Sheet} The sheet for the specified week
*/
function getWeekSheet(date) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const weekStart = new Date(date);
    weekStart.setDate(date.getDate() - date.getDay()); // Start of week (Sunday)

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const weekSheetName = `Papers Week of ${months[weekStart.getMonth()]} ${weekStart.getDate()}`;

    // Try to get existing sheet
    let sheet = spreadsheet.getSheetByName(weekSheetName);

    // If sheet doesn't exist, create it
    if (!sheet) {
        sheet = spreadsheet.insertSheet(weekSheetName);
        // Add headers
        sheet.getRange(1, 1, 1, 5).setValues([['Type', 'Title', 'Author', 'Link', 'Date']]);
        sheet.getRange(1, 1, 1, 5).setFontWeight('bold');

        // Set column widths
        sheet.setColumnWidth(1, 120); // Type
        sheet.setColumnWidth(2, 400); // Title
        sheet.setColumnWidth(3, 200); // Author
        sheet.setColumnWidth(4, 150); // Link
        sheet.setColumnWidth(5, 100); // Date
    }

    return sheet;
}

/**
* Writes paper data to the current week's sheet
* @param {Array} papers - Array of paper objects containing title, author, link, and date
*/
function writePapersToSheet(papers) {
    try {
        // Get the date from the first paper or use current date
        const date = papers.length > 0 && papers[0].date ? papers[0].date : new Date();
        const sheet = getWeekSheet(date);

        // Get the last row with data
        const lastRow = sheet.getLastRow();
        const startRow = lastRow > 0 ? lastRow + 1 : 2; // Start after headers if sheet is empty

        // Prepare data
        const paperData = papers.map(paper => [
            paper.type,
            paper.title,
            paper.author,
            paper.link,
            paper.date ? paper.date.toLocaleDateString() : ''
        ]);

        // Write all data at once
        if (paperData.length > 0) {
            const range = sheet.getRange(startRow, 1, paperData.length, 5);
            range.setValues(paperData);

            // Batch process hyperlinks
            const formulas = paperData.map(row => [`=HYPERLINK("${row[3]}", "Open Link")`]);
            sheet.getRange(startRow, 4, formulas.length, 1).setFormulas(formulas);
            sheet.getRange(startRow, 4, formulas.length, 1).setWrap(true);
        }

        Logger.log(`Updated sheet "${sheet.getName()}" with ${paperData.length} new papers`);

    } catch (e) {
        Logger.log(`Error updating sheet: ${e.message}`);
        throw e;
    }
}

/**
* Deduplicates papers based on normalized title and concatenates authors for duplicate titles
* @param {Array} papers - Array of paper objects containing title, author, and type
* @return {Array} Deduplicated array of papers
*/
function deduplicatePapers(papers) {
    // Create a Map for faster lookups
    const paperMap = new Map();
    const typePriority = { 'new_article': 1, 'citation': 2, 'related_research': 3 };

    // Single pass through the array
    for (const paper of papers) {
        const normalizedTitle = normalizeTitleForComparison(paper.title);
        const existing = paperMap.get(normalizedTitle);

        if (existing) {
            // Update type if new paper has higher priority
            if (paper.type && existing.type &&
                typePriority[paper.type] < typePriority[existing.type]) {
                existing.type = paper.type;
            }
            // Append author if not already present
            if (!existing.author.includes(paper.author)) {
                existing.author = `${existing.author}, ${paper.author}`;
            }
        } else {
            paperMap.set(normalizedTitle, { ...paper });
        }
    }

    return Array.from(paperMap.values());
}

/**
* Extracts the actual paper URL from a Google Scholar redirect URL
* @param {string} scholarUrl - The Google Scholar redirect URL
* @return {string} The actual paper URL
*/
function extractActualUrl(scholarUrl) {
    try {
        // Extract the URL parameter from the scholar_url
        const urlMatch = scholarUrl.match(/[?&]url=([^&]+)/i);
        if (!urlMatch) return scholarUrl;

        // Decode the URL twice because Google Scholar double-encodes some URLs
        let decodedUrl = decodeURIComponent(urlMatch[1]);
        try {
            // Try second decode for doubly-encoded URLs
            decodedUrl = decodeURIComponent(decodedUrl);
        } catch (e) {
            // If second decode fails, use the first decode result
        }

        return decodedUrl;
    } catch (e) {
        Logger.log(`Error extracting actual URL: ${e.message}`);
        return scholarUrl;
    }
}

/**
* Extracts author name from subject line
* @param {string} subject - The email subject line
* @return {string} The extracted author name
*/
function extractAuthorFromSubject(subject) {
    // Pattern 1: "X new citations to articles by Author Name"
    const citationPattern = /(\d+)\s+new\s+citations\s+to\s+articles\s+by\s+(.*?)(?:\s+-|$)/i;

    // Pattern 2: "Author Name - new related research"
    const relatedPattern = /^(.*?)\s+-\s+new\s+related\s+research/i;

    // Pattern 3: "Author Name - new articles"
    const newArticlePattern = /^(.*?)\s+-\s+new\s+articles/i;

    let author = '';

    // Try each pattern in order
    const citationMatch = subject.match(citationPattern);
    if (citationMatch) {
        author = citationMatch[2].trim();
    } else {
        const relatedMatch = subject.match(relatedPattern);
        if (relatedMatch) {
            author = relatedMatch[1].trim();
        } else {
            const newArticleMatch = subject.match(newArticlePattern);
            if (newArticleMatch) {
                author = newArticleMatch[1].trim();
            }
        }
    }

    return author;
}

/**
* Gets papers from emails within a date range
* @param {Date} fromDate - Start date for email search
* @param {Date} toDate - End date for email search
* @return {Array} Array of paper objects
*/
function getPapersFromDateRange(fromDate, toDate) {
    const IGNORED_URLS = new Set([
        'scholar.google.com/scholar_alerts',
        'scholar.google.com/scholar_settings',
        'support.google.com',
        'google.com/intl/',
        'mail.google.com',
        'accounts.google.com'
    ]);

    // Format dates for Gmail search
    const formatDate = (date) => {
        return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
    };

    const query = `from:scholaralerts-noreply@google.com after:${formatDate(fromDate)} before:${formatDate(toDate)}`;
    const threads = GmailApp.search(query);

    if (!threads.length) {
        Logger.log(`No Scholar Alert emails found between ${formatDate(fromDate)} and ${formatDate(toDate)}.`);
        return [];
    }

    // Pre-compile regex patterns
    const linkPattern = /<a[^>]+href=["']([^"']+)["'][^>]*>([^<]+)<\/a>/gi;
    const contentEndPattern = /This message was sent by Google Scholar/;

    const listOfLinks = [];

    for (let i = 0; i < threads.length; i++) {
        const thread = threads[i];
        const subject = thread.getFirstMessageSubject();
        const latestMessage = thread.getMessages().pop();
        const author = extractAuthorFromSubject(subject);
        const paperType = getType(subject);

        try {
            const htmlBody = latestMessage.getBody();
            const contentEndIndex = htmlBody.search(contentEndPattern);
            const contentToProcess = contentEndIndex !== -1 ?
                htmlBody.substring(0, contentEndIndex) :
                htmlBody;

            let match;
            while ((match = linkPattern.exec(contentToProcess)) !== null) {
                const url = match[1];
                const text = match[2].trim();

                if (url.includes('scholar.google.com') &&
                    !Array.from(IGNORED_URLS).some(ignored => url.includes(ignored))) {
                    listOfLinks.push({
                        title: text,
                        link: extractActualUrl(url),
                        author: author,
                        date: latestMessage.getDate(),
                        type: paperType
                    });
                }
            }

            thread.markRead();

        } catch (e) {
            Logger.log(`Error processing thread ${i + 1}: ${e.message}`);
        }
    }

    return listOfLinks;
}

/**
* Gets papers from the last week's emails
* @return {Array} Array of paper objects
*/
function getPapersFromLastWeek() {
    const toDate = new Date();
    const fromDate = new Date();
    fromDate.setDate(fromDate.getDate() - 7);
    return getPapersFromDateRange(fromDate, toDate);
}

/**
* Normalizes a title for comparison by removing extra spaces, special characters, and converting to lowercase
* @param {string} title - The title to normalize
* @return {string} Normalized title for comparison
*/
function normalizeTitleForComparison(title) {
    return title.toString()
        .toLowerCase()
        .replace(/\s+/g, ' ')  // Replace multiple spaces with single space
        .replace(/[^\w\s]/g, '') // Remove special characters
        .trim();
}

/**
* Gets existing papers from all sheets
* @return {Map} Map of normalized paper titles to paper objects
*/
function getExistingPapers() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const existingPapers = new Map();

    for (const sheet of sheets) {
        const data = sheet.getDataRange().getValues();
        // Skip header row
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row[1]) { // Title column
                const originalTitle = row[1].toString();
                const normalizedTitle = normalizeTitleForComparison(originalTitle);
                existingPapers.set(normalizedTitle, {
                    title: originalTitle, // Keep original title for display
                    author: row[2],
                    type: row[0]
                });
            }
        }
    }

    return existingPapers;
}

/**
* Main function to process ResearchGate and Google Scholar emails
* and add paper details to a Google Sheet.
* This function will be triggered daily.
*/
function getPapersFromGmail() {
    // Get papers from last week
    const newPapers = getPapersFromLastWeek();
    if (!newPapers.length) {
        Logger.log("No new papers found.");
        return [];
    }

    // Get existing papers
    const existingPapers = getExistingPapers();

    // Filter out duplicates and merge authors for existing papers
    const uniqueNewPapers = newPapers.filter(paper => {
        const normalizedTitle = normalizeTitleForComparison(paper.title);
        const existing = existingPapers.get(normalizedTitle);
        if (existing) {
            return false;
        }
        return true;
    });

    // Deduplicate remaining papers
    const finalPapers = deduplicatePapers(uniqueNewPapers);

    Logger.log(`Found ${finalPapers.length} new unique papers from ${newPapers.length} total papers`);

    // Write to sheet if we have results
    if (finalPapers.length > 0) {
        writePapersToSheet(finalPapers);
    }

    return finalPapers;
}

/**
* Process papers from a specific date range
* @param {Date} fromDate - Start date for email search
* @param {Date} toDate - End date for email search
*/
function processDateRange(fromDate, toDate) {
    const papers = getPapersFromDateRange(fromDate, toDate);
    if (!papers.length) {
        Logger.log(`No papers found between ${fromDate.toLocaleDateString()} and ${toDate.toLocaleDateString()}`);
        return;
    }

    // Get existing papers
    const existingPapers = getExistingPapers();

    // Define type priority
    const typePriority = { 'new_article': 1, 'citation': 2, 'related_research': 3 };

    // Filter out duplicates and merge authors for existing papers
    const uniqueNewPapers = papers.filter(paper => {
        const existing = existingPapers.get(paper.link);
        if (existing) {
            // If paper exists, update author list if needed
            if (!existing.author.includes(paper.author)) {
                existing.author = `${existing.author}, ${paper.author}`;
            }
            // Update type if new paper has higher priority
            if (paper.type && existing.type &&
                typePriority[paper.type] < typePriority[existing.type]) {
                existing.type = paper.type;
            }
            return false;
        }
        return true;
    });

    // Deduplicate remaining papers
    const finalPapers = deduplicatePapers(uniqueNewPapers);

    Logger.log(`Found ${finalPapers.length} new unique papers from ${papers.length} total papers`);

    // Write to sheet if we have results
    if (finalPapers.length > 0) {
        writePapersToSheet(finalPapers);
    }
}

function getType(subject) {
    if (subject.includes('new citations to articles')) {
        return 'citations'
    }

    if (subject.includes('new related research')) {
        return 'related_research'
    }

    if (subject.includes('new articles')) {
        return 'new_article'
    }

    return 'new_article'
}

/**
* Process papers from April 2025
* This function will create weekly sheets and process all papers from April 1st, 2025
*/
function processFromApril2025() {
    // Set start date to April 1st, 2025
    const fromDate = new Date(2025, 4, 25); // Month is 0-based, so 3 = April
    const toDate = new Date();

    // Get all existing papers at the start to maintain global state
    const existingPapers = getExistingPapers();
    Logger.log(`Found ${existingPapers.size} existing papers across all sheets`);

    // Process in weekly chunks to avoid timeouts
    let currentFrom = fromDate;
    while (currentFrom < toDate) {
        let currentTo = new Date(currentFrom);
        currentTo.setDate(currentTo.getDate() + 6); // Add 6 days to get a week
        if (currentTo > toDate) currentTo = toDate;

        Logger.log(`Processing papers from ${currentFrom.toLocaleDateString()} to ${currentTo.toLocaleDateString()}`);

        // Get papers for this week
        const papers = getPapersFromDateRange(currentFrom, currentTo);
        if (papers.length > 0) {
            // Filter out duplicates using the global existing papers map
            const uniqueNewPapers = papers.filter(paper => {
                const existing = existingPapers.get(paper.link);
                if (existing) {
                    // If paper exists, update author list if needed
                    if (!existing.author.includes(paper.author)) {
                        existing.author = `${existing.author}, ${paper.author}`;
                    }
                    // Update type if new paper has higher priority
                    const typePriority = { 'new_article': 1, 'citation': 2, 'related_research': 3 };
                    if (paper.type && existing.type &&
                        typePriority[paper.type] < typePriority[existing.type]) {
                        existing.type = paper.type;
                    }
                    return false;
                }
                // Add to existing papers map
                existingPapers.set(paper.link, { ...paper });
                return true;
            });

            // Deduplicate remaining papers
            const finalPapers = deduplicatePapers(uniqueNewPapers);

            Logger.log(`Found ${finalPapers.length} new unique papers from ${papers.length} total papers for this week`);

            // Write to sheet if we have results
            if (finalPapers.length > 0) {
                writePapersToSheet(finalPapers);
            }
        }

        // Move to next week
        currentFrom = new Date(currentTo);
        currentFrom.setDate(currentFrom.getDate() + 1);

        // Add a small delay between weeks to avoid rate limits
        Utilities.sleep(1000);
    }
}