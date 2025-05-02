// Mock for pdf-parse
const pdfParse = jest.fn().mockResolvedValue({
  text: 'Mocked PDF content',
  numpages: 1,
  info: {
    Title: 'Mock PDF',
    Author: 'Jest',
    Creator: 'Test Suite'
  }
});

module.exports = pdfParse;
