| case | input | input description                         | output         | output description                                     |
|------|-------|------------------------------------------|----------------------|--------------------------------------------------------|
| 1    | October-December, 2001                     | Month1 to Month2 in Year | 2001-10/2001-12 | Year-Padded Month1/Year-Padded Month2                 |
| 2    | January 24, 2014 - February 24, 2018        | Month1 Day1, Year1 - Month2 Day2, Year2 | 2014-01-24/2018-02-24 | Year1-Padded Month1-Padded Day1/Year2-Padded Month2-Padded Day|
| 3    | undated                                     | String Undated          | 0000/0000       | String 0000/0000                                      |
| 4    | 1958-1986 and undated                      | Some Year Range and Undated | 1958/1986   | The year range                                         |
| 5    | c 1790s                                    | The decade that the year is in | 1790/1799 | the year range for the decade                         |
| 6    | 1790s                                      | decade that the year is in | 1790/1799      | the year range for the decade                         |
| 7    | 1970s-1980s                                 | decades used start starting and ending points | 1970/1989 | starting year of decade range / ending year of decade range |
| 8    | October, 2001                              | Month and Year          | 2001-10         | year-padded month                                      |
| 9    | 1978-1984                                  | year range              | 1978/1984       | start year / end  year                                 |
| 10   | c. 1978                                    | around year             | 1978            | the year                                               |
| 11   | Spring, 2001                               | season and year         | 2001            | the year                                               |
| 12   | October 16, 2001                           | Month day year          | 2001-10-16      | year-padded month-padded day                           |
| 13   | October 16-18, 2001                        | Month day range year    | 2001-10-16/2001-10-18 | year-padded month-padded day/year-padded month-padded day |
| 14   | c. 1945-1947                               | circa year range        | 1945/1947       | Year1/Year2                                            |
| 15   | circa 1945                                 | circa year              | 1945            | year                                                   |
| 16   | c. 1946                                    | circa year              | 1946            | year                                                   |
| 17   | 1942, 1045, 1945-1947                      | Year1, Year2, Year Range | 1045/1947    | Consider all years, determine start and end year, present start year/end year |
| 18   | 1942                                       | year                    | 1942            | year                                                   |
