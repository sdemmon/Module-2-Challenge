# Module-2-Challenge
Challenge
In this challenge we were tasked to compile the yearly change in stock price, percentage change in stock price and the stock volume from daily data of a particular stock. This project showcases a meticulously crafted VBA script aimed at streamlining the process of analyzing stock data in an Excel workbook. The sequence of commands in the script highlights its effectiveness in delivering accurate insights from complex stock datasets across multiple years.

The script's architecture begins with Looping Through Worksheets, a pivotal step that enables the analysis to span various years effortlessly. This iteration is achieved through the For Each loop, where the script iterates through each worksheet, extracting and processing the data for analysis.

Each worksheet undergoes a Summary Table Setup phase, facilitated by the use of ws.Cells(row, column).Value. This command strategically places the header labels in the designated cells of the worksheet. It structures the layout for the subsequent analysis, enhancing clarity and organization.

The Loop Through Rows section capitalizes on the loop's capability to traverse individual rows within each worksheet. This iterative movement is vital for calculating various metrics. The script dynamically calculates values such as Yearly Change, Percentage Change, and Total Stock Volume by employing mathematical operations and referencing specific cell values. The conditional formatting, achieved through Interior.Color changes, visually demarcates positive and negative yearly changes, aiding in quick trend identification.

The script's proficiency shines in the Identifying Extreme Values phase. Here, it accurately identifies stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The use of WorksheetFunction.Max() and WorksheetFunction.Min() allows the script to find these extremes within the dataset dynamically.

Closing the loop, the script employs Next to transition to the next worksheet, ensuring a comprehensive analysis across all years of data.

In essence, this VBA script’s intricately ordered sequence of commands transforms the manual and error-prone process of stock analysis into an automated, precise, and efficient endeavor. It empowers decision-makers by quickly revealing trends and outliers, thus serving as an indispensable tool for making informed choices in the dynamic realm of stock trading.
