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
* Writes paper data to a new sheet named with current date
* @param {Array} papers - Array of paper objects containing title, author, link, and date
*/
function writePapersToSheet(papers) {
    // Get the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Create a new sheet with today's date
    const newSheetName = createSheetName();
    let sheet;

    try {
        // Try to get existing sheet with same name
        sheet = spreadsheet.getSheetByName(newSheetName);

        if (sheet) {
            // If sheet exists, add a number to make it unique
            let counter = 1;
            let uniqueName = `${newSheetName} (${counter})`;
            while (spreadsheet.getSheetByName(uniqueName)) {
                counter++;
                uniqueName = `${newSheetName} (${counter})`;
            }
            sheet = spreadsheet.insertSheet(uniqueName);
        } else {
            // Create new sheet if it doesn't exist
            sheet = spreadsheet.insertSheet(newSheetName);
        }

        // Move the new sheet to the front
        sheet.activate();
        spreadsheet.moveActiveSheet(1);

        // Prepare all data at once
        const headers = [['Title', 'Author', 'Link', 'Date']];
        const paperData = papers.map(paper => [
            paper.title,
            paper.author,
            paper.link,
            paper.date
        ]);

        // Write all data at once
        const range = sheet.getRange(1, 1, paperData.length + 1, 4);
        range.setValues([...headers, ...paperData]);

        // Batch formatting
        sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
        sheet.autoResizeColumns(1, 4);

        // Batch process hyperlinks
        const formulas = paperData.map(row => [`=HYPERLINK("${row[2]}", "Open Link")`]);
        if (formulas.length > 0) {
            sheet.getRange(2, 3, formulas.length, 1).setFormulas(formulas);
            sheet.getRange(2, 3, formulas.length, 1).setWrap(true);
        }

        // Add timestamp and count in the sheet
        const timestamp = new Date().toLocaleString();
        sheet.getRange(paperData.length + 3, 1, 1, 2).setValues([[`Last Updated:`, timestamp]]);
        sheet.getRange(paperData.length + 4, 1, 1, 2).setValues([[`Total Papers:`, paperData.length]]);

        Logger.log(`Created new sheet "${sheet.getName()}" with ${paperData.length} papers`);

    } catch (e) {
        Logger.log(`Error creating/updating sheet: ${e.message}`);
        throw e;
    }
}

/**
* Deduplicates paper links and concatenates authors for duplicate URLs
* @param {Array} links - Array of paper links with title, link, and author
* @return {Array} Deduplicated array of paper links
*/
function deduplicatePaperLinks(links) {
    return Array.from(links.reduce((map, paper) => {
        const existing = map.get(paper.link);
        if (existing && !existing.author.includes(paper.author)) {
            existing.author = `${existing.author}, ${paper.author}`;
        } else if (!existing) {
            map.set(paper.link, paper);
        }
        return map;
    }, new Map()).values());
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
* Main function to process ResearchGate and Google Scholar emails
* and add paper details to a Google Sheet.
* This function will be triggered daily.
*/
function getPapersFromGmail() {
    // Cache frequently used values
    const IGNORED_URLS = new Set([
        'scholar.google.com/scholar_alerts',
        'scholar.google.com/scholar_settings',
        'support.google.com',
        'google.com/intl/',
        'mail.google.com',
        'accounts.google.com'
    ]);

    const query = 'from:scholaralerts-noreply@google.com';
    const threads = GmailApp.search(query);

    if (!threads.length) {
        Logger.log("No Scholar Alert emails found.");
        return [];
    }

    // Pre-compile regex pattern
    const linkPattern = /<a[^>]+href=["']([^"']+)["'][^>]*>([^<]+)<\/a>/gi;
    const authorPattern = /by\s+(.+?)(?:\s+-|$)/i;

    // Process all threads in bulk
    const listOfLinks = threads.reduce((allLinks, thread, index) => {
        const subject = thread.getFirstMessageSubject();
        const latestMessage = thread.getMessages().pop();

        // Extract author more efficiently
        const authorMatch = subject.match(authorPattern);
        const author = authorMatch ? authorMatch[1].trim() : '';

        Logger.log(`Processing thread ${index + 1}/${threads.length}`);

        try {
            const htmlBody = latestMessage.getBody();
            const contentEndIndex = htmlBody.indexOf('This message was sent by Google Scholar');
            const contentToProcess = contentEndIndex !== -1 ?
                htmlBody.substring(0, contentEndIndex) :
                htmlBody;

            // Process all links in one go
            const matches = [...contentToProcess.matchAll(linkPattern)];
            const validLinks = matches
                .map(match => ({
                    url: match[1],
                    text: match[2].trim()
                }))
                .filter(({ url }) =>
                    url.includes('scholar.google.com') &&
                    !Array.from(IGNORED_URLS).some(ignored => url.includes(ignored))
                )
                .map(({ url, text }) => ({
                    title: text,
                    link: extractActualUrl(url),
                    author: author,
                    date: latestMessage.getDate()
                }));

            allLinks.push(...validLinks);

        } catch (e) {
            Logger.log(`Error processing thread ${index + 1}: ${e.message}`);
        }

        return allLinks;
    }, []);

    // Process results
    const uniqueLinks = deduplicatePaperLinks(listOfLinks);
    Logger.log(`Found ${uniqueLinks.length} unique papers from ${threads.length} threads`);

    // Write to sheet if we have results
    if (uniqueLinks.length > 0) {
        writePapersToSheet(uniqueLinks);
    }

    return uniqueLinks;
}





https://scholar.google.com/scholar_url?url=https://pubs.aip.org/aip/acp/article-abstract/3291/1/020005/3348459&amp;hl=en&amp;sa=X&amp;d=14263522341197601930&amp;ei=fuREaL7_MfiJ6rQPkqSymQY&amp;scisig=AAZF9b-nroaN49BdNNbtQCEZI9qt&amp;oi=scholaralrt&amp;hist=6X-hv9AAAAAJ:15873594094838026202:AAZF9b9uu5fG2B59mEjL8GPhe_qf&amp;html=&amp;pos=3&amp;folt=cit



<a href="https://scholar.google.com/scholar_url?url=https://arxiv.org/pdf/2506.03283&amp;hl=en&amp;sa=X&amp;d=200998896991037433&amp;ei=fuREaL7_MfiJ6rQPkqSymQY&amp;scisig=AAZF9b_9W2A-NQAXxkYzgRHwtFh8&amp;oi=scholaralrt&amp;hist=6X-hv9AAAAAJ:15873594094838026202:AAZF9b9uu5fG2B59mEjL8GPhe_qf&amp;html=&amp;pos=0&amp;folt=cit" class="m_5785777122273086187gse_alrt_title" style="font-size:17px;color:#1a0dab;line-height:22px" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://scholar.google.com/scholar_url?url%3Dhttps://arxiv.org/pdf/2506.03283%26hl%3Den%26sa%3DX%26d%3D200998896991037433%26ei%3DfuREaL7_MfiJ6rQPkqSymQY%26scisig%3DAAZF9b_9W2A-NQAXxkYzgRHwtFh8%26oi%3Dscholaralrt%26hist%3D6X-hv9AAAAAJ:15873594094838026202:AAZF9b9uu5fG2B59mEjL8GPhe_qf%26html%3D%26pos%3D0%26folt%3Dcit&amp;source=gmail&amp;ust=1749542361707000&amp;usg=AOvVaw3-UVSRCEBvtFKv1skh6dOz"><span class="il">Empirical</span> <span class="il">Evaluation</span> of <span class="il">Generalizable</span> <span class="il">Automated</span> <span class="il">Program</span> <span class="il">Repair</span> with <span class="il">Large</span> <span class="il">Language</span> <span class="il">Models</span></a>



https://scholar.google.com/scholar_url?url=https://search.ebscohost.com/login.aspx%3Fdirect%3Dtrue%26profile%3Dehost%26scope%3Dsite%26authtype%3Dcrawler%26jrnl%3D2158107X%26AN%3D185173680%26h%3DGtg6zNQzIEuGME7aLkKd6ioGLJ6CTLJHop4ueDJyFeEdu3ct21MY3zRzgSUvCXD3YAFb7oHGCeTfnvwLKFZWYw%253D%253D%26crl%3Dc&amp;hl=en&amp;sa=X&amp;d=7553936511393527748&amp;ei=fuREaL7_MfiJ6rQPkqSymQY&amp;scisig=AAZF9b-HI7u_hJHTA4S11i_f7eJr&amp;oi=scholaralrt&amp;hist=6X-hv9AAAAAJ:15873594094838026202:AAZF9b9uu5fG2B59mEjL8GPhe_qf&amp;html=&amp;pos=1&amp;folt=cit


<a href="https://scholar.google.com/scholar_url?url=https://search.ebscohost.com/login.aspx%3Fdirect%3Dtrue%26profile%3Dehost%26scope%3Dsite%26authtype%3Dcrawler%26jrnl%3D2158107X%26AN%3D185173680%26h%3DGtg6zNQzIEuGME7aLkKd6ioGLJ6CTLJHop4ueDJyFeEdu3ct21MY3zRzgSUvCXD3YAFb7oHGCeTfnvwLKFZWYw%253D%253D%26crl%3Dc&amp;hl=en&amp;sa=X&amp;d=7553936511393527748&amp;ei=fuREaL7_MfiJ6rQPkqSymQY&amp;scisig=AAZF9b-HI7u_hJHTA4S11i_f7eJr&amp;oi=scholaralrt&amp;hist=6X-hv9AAAAAJ:15873594094838026202:AAZF9b9uu5fG2B59mEjL8GPhe_qf&amp;html=&amp;pos=1&amp;folt=cit" class="m_5785777122273086187gse_alrt_title" style="font-size:17px;color:#1a0dab;line-height:22px" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://scholar.google.com/scholar_url?url%3Dhttps://search.ebscohost.com/login.aspx%253Fdirect%253Dtrue%2526profile%253Dehost%2526scope%253Dsite%2526authtype%253Dcrawler%2526jrnl%253D2158107X%2526AN%253D185173680%2526h%253DGtg6zNQzIEuGME7aLkKd6ioGLJ6CTLJHop4ueDJyFeEdu3ct21MY3zRzgSUvCXD3YAFb7oHGCeTfnvwLKFZWYw%25253D%25253D%2526crl%253Dc%26hl%3Den%26sa%3DX%26d%3D7553936511393527748%26ei%3DfuREaL7_MfiJ6rQPkqSymQY%26scisig%3DAAZF9b-HI7u_hJHTA4S11i_f7eJr%26oi%3Dscholaralrt%26hist%3D6X-hv9AAAAAJ:15873594094838026202:AAZF9b9uu5fG2B59mEjL8GPhe_qf%26html%3D%26pos%3D1%26folt%3Dcit&amp;source=gmail&amp;ust=1749542361708000&amp;usg=AOvVaw3lJl-r_jZ9Xu3xpshNNV46">From Code Analysis to Fault Localization: A Survey of Graph Neural Network Applications in Software Engineering.</a>