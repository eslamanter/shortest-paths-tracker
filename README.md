# Shortest Paths Tracker 2022.1 | Description	 

# What is it for? 

Shortest Paths Tracker is an executable application designed to analyze directed graphs and find the shortest paths in terms of total cost for all connected pairs of nodes. 

# How to run? 

Shortest Paths Tracker 2022.1 has two input modes: read and write modes. Read mode imports graph information from an [.xlsx] excel workbook file. The file must exist in the same directory of the application and can’t be contemporarily opened in excel while running the analysis. The workbook must be already saved on an active worksheet that contains the graph information in the first four columns representing respectively tail node, head node, arc cost, and number of directions. First row should contain the titles and can’t be left blank. Values must be inserted starting from the second row. Direction of arcs matters, and each arc can be defined once through unique tail/head pair. Tail/head nodes can only assume positive integer values, while costs can also assume decimal numbers. Arcs can be 1-way by assigning 1 or leaving the cell blank or 2-way by assigning 2 to duplicate the arc in the opposite direction keeping the same cost. Write mode receives the same information assigned manually arc by arc. In either mode, the name of the workbook is required to write the results. The results of the analysis if run successfully are two worksheets added to the workbook and a text file of the same name created in the same directory. The first worksheet contains a forward star representation of the graph. The second worksheet contains the shortest path for every connected pair of nodes and the corresponding cost. The text file contains run date, time, and duration in addition to possible warnings and errors. 

# How it works? 

Shortest Paths Tracker 2022.1 firstly performs a sequence of verifications for the validity of the file, active worksheet, graph information related to the values of nodes, costs, and directions, and logs any incorrect input to the text file. When all are valid, it creates and writes the forward star of the graph where arcs are sorted in ascending order primarily by tail nodes and secondarily by head ones. Starting from the forward star, it applies Dijkstra’s shortest path algorithm to find the paths of minimum costs connecting an origin node with every other node in the graph. This algorithm terminates after a maximum of n iterations in addition to the initialization for a graph of n nodes. The algorithm is applied to all nodes iterating the origin node every time. The resulting number of all feasible shortest paths may reach n2 – n in a completely connected graph. A maximum of 1000 nodes per graph has been set for issues of time complexity O(n2) and limitations of excel worksheet size. 

2022 Eslam Anter
