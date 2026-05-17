# Movie Dashboard
## Table of Content
- [Problem Statement](#problem-statement)
- [Data Source](#data-source)
- [Tools](#tools)
- [Data Cleaning and Preparation](#data-cleaning-and-preparation)
- [Recommendations](#Recommendations)
- [M Code](#m-code)
### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres.

Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify:

- The best-performing actors
- The top movies based on box office metrics
- Director performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions

The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.
### Data Source

Movie Data:
The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie’s performance (box office and budget), actors, directors, etc.
Find Original Data Source Here: [Movie Dataset Excel file](https://github.com/user-attachments/files/27907274/Movies_Data_Homework.xlsx)
### Tools
1. Power Query – Used Power Query for data cleaning
2. Excel – Used Excel for data analysis
3. Pivot Tables – For creating the dashboard and visualizations
### Data Cleaning and Preparation
- Data loading and inspection.
- Handling errors, missing values.
- Data cleaning and formatting.
- 
The Excel file after the data cleaning & preparation process can be downloaded here – [Ready to use Dashboard](https://github.com/user-attachments/files/27935360/Movies_Data_Homework_Alena_Bondar_dashoard.xlsx)
### Dashboard
<img width="1203" height="669" alt="image" src="https://github.com/user-attachments/assets/0016a005-8a44-450f-ae6a-a3f9963526e2" />
<img width="1223" height="797" alt="image" src="https://github.com/user-attachments/assets/198e09f5-1a9e-45ae-b5fa-70fa17f0aa06" />
### M Code

```
let
    Source = Excel.Workbook(File.Contents("C:\Users\User\Downloads\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Removed Blank Rows" = Table.SelectRows(#"Changed Type", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Blank Rows",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Merged Queries", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Genres",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries1" = Table.NestedJoin(#"Reordered Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genres.Genre.1"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Expanded Genres1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.1", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns1",{{"Genres.Genre.1", "Genres 2"}}),
    #"Merged Queries2" = Table.NestedJoin(#"Renamed Columns", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries2", "Directors", {"Director"}, {"Directors.Director"}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Expanded Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries3" = Table.NestedJoin(#"Reordered Columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries3", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Expanded Actors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries4" = Table.NestedJoin(#"Reordered Columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors1" = Table.ExpandTableColumn(#"Merged Queries4", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Expanded Actors1",{{"Actors.Actor.1", "Actors 2"}}),
    #"Reordered Columns4" = Table.ReorderColumns(#"Renamed Columns1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors 2", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries5" = Table.NestedJoin(#"Reordered Columns4", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors2" = Table.ExpandTableColumn(#"Merged Queries5", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns5" = Table.ReorderColumns(#"Expanded Actors2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors 2", "Cast_Third_ID", "Actors.Actor.1", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns5",{{"Actors.Actor.1", "Actors 3"}}),
    #"Merged Queries6" = Table.NestedJoin(#"Renamed Columns2", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors3" = Table.ExpandTableColumn(#"Merged Queries6", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns6" = Table.ReorderColumns(#"Expanded Actors3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors 2", "Cast_Third_ID", "Actors 3", "Cast_Fourth_ID", "Actors.Actor.1", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Reordered Columns6",{{"Actors.Actor.1", "Actors 4"}}),
    #"Merged Queries7" = Table.NestedJoin(#"Renamed Columns3", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors4" = Table.ExpandTableColumn(#"Merged Queries7", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns7" = Table.ReorderColumns(#"Expanded Actors4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors 2", "Cast_Third_ID", "Actors 3", "Cast_Fourth_ID", "Actors 4", "Cast_Fifth_ID", "Actors.Actor.1", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Reordered Columns7",{{"Actors.Actor.1", "Actors 5"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns4", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"ROI", Percentage.Type}, {"Box Office Revenue ($)", Currency.Type}, {"Budget ($)", Currency.Type}}),
    #"Added Custom1" = Table.AddColumn(#"Changed Type1", "Profit", each [#"Box Office Revenue ($)"]-[#"Budget ($)"]),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom1",{{"Profit", Currency.Type}})
in
    #"Changed Type2"
```
### Recommendations
Recommendation on the Next Movie Production Strategy:
The decision on what type of movie to produce next should be based on the company’s current goals and financial capacity.

Based on the analysis of data from 2012 to 2016, the action genre consistently generates the highest total revenue and profit. Although these movies require larger budgets, they also bring the biggest returns. This makes action films a strong choice if the company is ready to invest more in order to achieve maximum profit.

From a timing perspective, the data also shows clear seasonal patterns. For action movies, the most successful release periods are April and the summer months from May to July, which consistently generate the highest box office results. Releasing films during these periods can further maximize revenue potential.

From a casting perspective, Tom Cruise stands out as the top-performing actor in the action genre. Over the five-year period, he appeared in five action movies and generated around $1.39 billion in total profit. Another strong performer is Hugh Jackman, who contributed approximately $842.7 million across two films. These actors can be considered reliable choices for leading roles in high-budget action productions.

Looking at directors, Zack Snyder delivered the strongest results, generating over $1 billion in profit from two films. Bryan Singer also showed strong performance, with around $914 million in total profit. Both directors have proven their ability to deliver commercially successful action movies.

At the same time, if the company prefers a lower-budget approach or wants to reduce financial risk, the horror genre offers a strong alternative. Horror movies consistently show the highest return on investment (ROI), meaning they can generate solid returns with relatively small budgets.

Seasonality also plays an important role for this genre. Horror movies perform best when released in June, September, and January, indicating different audience behavior compared to action films and highlighting opportunities for strategic scheduling.

In this category, Patrick Wilson and Annabelle Wallis delivered the best results among actors. Among directors, James Wan stands out with approximately $481 million in profit, followed by John R. Leonetti with around $256 million.

Overall, the data suggests that the company should not rely on just one approach. A balanced strategy could be the most effective: investing in big action movies to drive high revenue, while also producing lower-budget horror films to ensure steady and efficient returns. This approach supports more informed, data-driven decision-making and helps balance growth with risk management.
