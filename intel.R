# Load the necessary libraries
library(readxl)      # For reading Excel files
library(tidyverse)   # For data manipulation and visualisation
library(cluster)     # For clustering methods
library(openxlsx)    # For exporting data to Excel

intel <- read_excel(file.choose()) # select and read the required file
# View Data
view(intel) # view the dataframe

# DATA EXPLORATION
# Display column names
names(intel) # display column names
# Display summary Statistics
summary(intel) # display summary statistics of all the columns


dfz <- scale(intel) # standardize the data to mean 0 and sd 1
# view the standardized data
view(dfz)

#################################SEGMENTATION###########################

# Calculate distance between observations of dataset
distance <- dist(dfz, method = 'euclidean')

# Hierarchical Clustering
hcw = hclust(distance, method = 'ward.D2') # use ward.D2 linkage methhod

# Plot the dendrogram to visualise the clustering
plot(hcw, main = "Cluster Dendrogram", xlab = "Observations", ylab = "Height")

# Determine the optimum number of clusters using the Elbow Plot
x <- c(1:10)
sort_height <- sort(hcw$height, decreasing = TRUE)

y <- sort_height[1:10]
# Plot elbow plot
plot(x, y, type = "b", main = "Elbow Plot for Ward.D2 method", xlab = "Number of Clusters", ylab = "Height")
lines(x, y, col = "blue")

# Display clusters on Dendogram
plot(hcw, main = "Cluster Dendrogram", xlab = "Observations", ylab = "Height")
rect.hclust(hcw, k = 3, border = 2:5)

# CUT DENDROGRAM INTO 3 CLUSTERS
# "cutree()" assigns each observation to a cluster.
cluster <- cutree(hcw, k = 3)  # Here dendrogram will be cut into 3 clusters
# Create a frequency table to see the size of each cluster
table(cluster)
# Add assigned clusters back to the original data
df_final <- cbind(intel, cluster)
# Check the updated dataset
View(df_final)

################### DESCRIPTION STEP ###################

# CALCULATE SEGMENT SIZES
# Proportions of each cluster
proportions <- table(df_final$cluster) / nrow(df_final)
percentages <- proportions * 100
# Display segment sizes in percentages
print(percentages)

# EXPLORE MEAN VALUES OF VARIABLES IN EACH CLUSTER
# Calculate mean values of selected variables grouped by cluster
segments<-
  df_final %>% 
  group_by(cluster) %>% 
  summarise(across(where(is.numeric), mean, .names = "{col}_mean"))
# Display the calculated means
segments


# SAVE MEAN TABLE TO EXCEL
# Export the summarised data to an Excel file
write.xlsx(segments, 'segments.xlsx')