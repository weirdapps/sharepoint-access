import { describe, expect, it } from 'vitest';

import { classifyLink, parseViewerPage, serverRelativeFromUrl } from '../src/sharepoint/links';

const TEAM = 'https://contoso.sharepoint.com';
const MY = 'https://contoso-my.sharepoint.com';
const GUID = '0f8fad5b-d9cb-469f-a165-70867728950e';

describe('serverRelativeFromUrl', () => {
  it('takes the path and drops the query', () => {
    expect(serverRelativeFromUrl(`${TEAM}/sites/team/Shared%20Documents/plan.docx?web=1`)).toBe(
      '/sites/team/Shared Documents/plan.docx',
    );
  });

  it('removes the r sharing prefix, which carries the server-relative path', () => {
    expect(serverRelativeFromUrl(`${TEAM}/:u:/r/sites/news/SitePages/Launch.aspx?csf=1&e=x`)).toBe(
      '/sites/news/SitePages/Launch.aspx',
    );
  });

  it('decodes Greek names', () => {
    expect(
      serverRelativeFromUrl(
        `${TEAM}/sites/t/Shared%20Documents/%CE%A3%CF%87%CE%AD%CE%B4%CE%B9%CE%BF.pptx`,
      ),
    ).toBe('/sites/t/Shared Documents/Σχέδιο.pptx');
  });
});

describe('classifyLink', () => {
  it.each([
    [`${MY}/:x:/g/personal/ann_contoso_com/EQabc?e=1`, 'file'],
    [`${TEAM}/:w:/s/team/EQdef`, 'file'],
    [`${TEAM}/:b:/r/sites/team/Shared%20Documents/report.pdf`, 'file'],
    [`${TEAM}/sites/team/Shared%20Documents/plan.docx`, 'file'],
    [`${TEAM}/sites/news/SitePages/Launch.aspx`, 'page'],
    [`${TEAM}/:u:/r/sites/news/SitePages/Launch.aspx?csf=1`, 'page'],
    [`${MY}/_layouts/15/onedrive.aspx`, 'other'],
    [`${TEAM}/_layouts/15/sharepoint.aspx`, 'other'],
    [`${TEAM}/:f:/g/sites/team/EQfolder`, 'other'],
    [`${MY}/:v:/g/personal/ann_contoso_com/EQvideo`, 'other'],
    [`${TEAM}/sites/team`, 'other'],
    [`${TEAM}/sites/team/Shared%20Documents/Forms/AllItems.aspx`, 'other'],
  ])('%s is %s', (url, kind) => {
    expect(classifyLink(url)).toBe(kind);
  });
});

describe('parseViewerPage', () => {
  const page = (context: string) =>
    `<html><script>var ctx = {${context}}; var x = {fileId:w.FileId,baseUrl:d.webAbsoluteUrl};</script></html>`;

  it('reads the file id, the web and the file name from the page context', () => {
    const html = page(
      `"FileId":"${GUID}","DocUrl":"https://x","FileName":"\\u03a3\\u03c7\\u03ad\\u03b4\\u03b9\\u03bf.xlsx","webAbsoluteUrlLegacy":null,"webAbsoluteUrl":"https://contoso-my.sharepoint.com/personal/ann_contoso_com"`,
    );
    expect(parseViewerPage(html)).toEqual({
      fileId: GUID,
      webAbsoluteUrl: 'https://contoso-my.sharepoint.com/personal/ann_contoso_com',
      fileName: 'Σχέδιο.xlsx',
    });
  });

  it('falls back to the sourcedoc parameter and unescapes the web', () => {
    const html =
      `<a href="\\u002f_layouts\\u002f15\\u002fDoc.aspx?sourcedoc={${GUID.toUpperCase()}}\\u0026action=default">x</a>` +
      page(`"webAbsoluteUrl":"https:\\/\\/contoso.sharepoint.com\\/sites\\/team"`);
    expect(parseViewerPage(html)).toEqual({
      fileId: GUID,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/team',
    });
  });

  it('returns null without a file id or without a web', () => {
    expect(
      parseViewerPage(page(`"webAbsoluteUrl":"https://contoso.sharepoint.com/sites/t"`)),
    ).toBeNull();
    expect(parseViewerPage(page(`"FileId":"${GUID}"`))).toBeNull();
    expect(parseViewerPage('<html><title>Sign in</title></html>')).toBeNull();
  });
});
