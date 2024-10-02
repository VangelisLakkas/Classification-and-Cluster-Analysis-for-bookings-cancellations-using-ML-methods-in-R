require(foreign)
require(xlsx)
library(readxl)
require(readxl)
library(lubridate)
require(ggplot2)
library(glmnet)
library(Matrix)
library(car)
library(Hmisc)
install.packages("HDclassif")
library('HDclassif')
install.packages("heplots")
library(heplots)
install.packages("caret")
library(caret)
library('nnet')
library(class)
library('e1071')
library('MASS')
install.packages("tree")
library('tree')
library(dplyr)
library(pROC)
install.packages("fpc")
library(fpc)





##Data Importing
bookings_df <-read_excel('project2.xlsx')
bookings_df$date.of.reservation <- ifelse(grepl("/", bookings_df$date.of.reservation),
                                          as.Date(bookings_df$date.of.reservation, format = "%m/%d/%Y"),
                                          as.Date(as.numeric(bookings_df$date.of.reservation), origin = "1899-12-30"))

bookings_df$date.of.reservation <- as.Date(bookings_df$date.of.reservation)
head(bookings_df$date.of.reservation)
length(unique(bookings_df$date.of.reservation))

class(bookings_df$date.of.reservation)

sum(is.na(bookings_df))
##After the conversion of the date of reservation column 2 NAs were produced since the number of NAs is very small compared to 
## the number of the observations I will use na.omit method to get rid of the NAs
bookings_df <- na.omit(bookings_df)
## Booking_ID is a unique identifier for each booking and not a feature that helps me 
##explain the cancelations of the bookings so I will remove it from the dataframe

bookings_df <- bookings_df[,-1]
bookings_df$month_of_reservation <- month(bookings_df$date.of.reservation)
bookings_df <- bookings_df[,-15]



bookings_df$type.of.meal <- as.factor(bookings_df$type.of.meal)
bookings_df$room.type <- as.factor(bookings_df$room.type)
bookings_df$market.segment.type <- as.factor(bookings_df$market.segment.type)
levels(bookings_df$market.segment.type)
bookings_df$car.parking.space <- as.factor(bookings_df$car.parking.space)
bookings_df$repeated <- as.factor(bookings_df$repeated)
bookings_df$month_of_reservation <- as.factor(bookings_df$month_of_reservation)
bookings_df$average.price <- as.numeric(bookings_df$average.price)


bookings_df$booking.status <- as.factor(bookings_df$booking.status)

summary(bookings_df$average.price)
quantile(bookings_df$average.price, probs = seq(0,1,0.25))
quantile(bookings_df$average.price, probs = seq(0,1,0.1))
head(table(bookings_df$average.price),20)

rows_tbclean <- bookings_df$average.price < 30
sum(rows_tbclean)
bookings_df$average.price[rows_tbclean] <- 30
head(table(bookings_df$average.price),20)


##variable selection
pairs(bookings_df[,c(1,2,3,4,5,6,7,8,9,11,12,13,14,16)],col=bookings_df$booking.status)
pairs(bookings_df[,c(1,2,3,4,5,6,7,9,11,12,13,14)],col=bookings_df$booking.status)
pairs(bookings_df[,c(1,2,3,5,6,7,9,11,14)],col=bookings_df$booking.status)

bookings_subset <- subset(bookings_df, select = c(1,2,3,5,6,7,9,11,14,15))

bookings_subset_bigger <- subset(bookings_df, select = c(1,2,3,5,6,7,9,11,13,14,15,16))


## test and train split
set.seed(123)
train_index <- createDataPartition(bookings_subset$booking.status, p = 0.75, list = FALSE)
train_index

training_subset <- bookings_subset[train_index, ]
testing_subset <- bookings_subset[-train_index, ]

train_index2 <- createDataPartition(bookings_df$booking.status, p = 0.75, list = FALSE)
train_index2

training_bookingsdf <- bookings_df[train_index2, ]
testing_bookingsdf <- bookings_df[-train_index2, ]

###AIC for variable selection for logistic regression
mfull <-  glm(booking.status~., data = training_bookingsdf, family = "binomial"(link="logit"))
summary(mfull)

step(mfull, trace = TRUE, direction = 'backward')

logistic_model <-  glm(booking.status~number.of.weekend.nights+car.parking.space+market.segment.type+lead.time+average.price+special.requests, data = training_bookingsdf, family = "binomial"(link="logit"))
summary(logistic_model)

logistic_preds <- predict(logistic_model, testing_bookingsdf, type ="response")
threshold <- 0.5  
bookings_preds <- ifelse(logistic_preds > threshold, 1, 0)
log_conf_matrix <- table(testing_bookingsdf$booking.status, bookings_preds)
log_conf_matrix

logistic_accuracy <- sum(diag(log_conf_matrix)) / sum(log_conf_matrix)
print(paste("Accuracy:", logistic_accuracy))


## checking for homogeneity
boxM(Y=bookings_subset[,c(1,2,3,8,9)],group=bookings_subset$booking.status) # homogeneity rejected
boxM(Y=bookings_df[,c(1,2,3,4,8,11,12,13,14)],group=bookings_df$booking.status) # homogeneity rejected

###### HDDA Model ######

hdda_model <- hdda(training_bookingsdf[,c(-5,-6,-7,-9,-10,-15,-16)], training_bookingsdf$booking.status, scaling = TRUE, model = "all", graph = TRUE)
hdda_pred <- predict(hdda_model, scale(testing_bookingsdf[,c(-5,-6,-7,-9,-10,-15,-16)]))
hdda_conf_matrix <- table(testing_bookingsdf$booking.status,hdda_pred$class)
print(hdda_conf_matrix)

hdda_accuracy <- sum(diag(hdda_conf_matrix)) / sum(hdda_conf_matrix)
print(paste("Accuracy:", hdda_accuracy))

##### Decision Tree #####
library(rpart)
tree_model <- rpart(booking.status ~ ., data = training_bookingsdf, method = "class")
tree_preds <- predict(tree_model, testing_bookingsdf[,c(-15)])
plot(tree_model)
summary(tree_model)
  text(tree_model)

install.packages('rpart.plot')
library(rpart.plot)
tree_conf_matrix <- table(testing_bookingsdf$booking.status, treepred_classes)
print(tree_conf_matrix)

insample_preds <- predict(tree_model, type = "class")
tree_conf_matrix_insample <- table(training_bookingsdf$booking.status, insample_preds)
insample_tree_accuracy <- sum(diag(tree_conf_matrix_insample)) / sum(tree_conf_matrix_insample)
insample_tree_accuracy

tree_accuracy <- sum(diag(tree_conf_matrix)) / sum(tree_conf_matrix)
print(paste("Accuracy:", tree_accuracy))
printcp(tree_model)


###prunning the tree

pruned_tree_model <- prune(tree_model, cp = 0.02) 
print(pruned_tree_model)
par(mfrow =c(2,1))

plot(pruned_tree_model)
summary(pruned_tree_model)
  text(pruned_tree_model)

rpart.plot(pruned_tree_model,type=5)
  
printcp(pruned_tree_model)
pruned_tree_preds <- predict(pruned_tree_model, testing_bookingsdf[, -c(15)], type = "class")
pruned_tree_conf_matrix <- table(testing_bookingsdf$booking.status, pruned_tree_preds)
pruned_tree_conf_matrix
pruned_tree_accuracy <- sum(diag(pruned_tree_conf_matrix)) / sum(pruned_tree_conf_matrix)
print(paste("Pruned Tree Accuracy:", pruned_tree_accuracy))

prunedtree_insample_preds <- predict(pruned_tree_model, type = "class")
prunedtree_conf_matrix_insample <- table(training_bookingsdf$booking.status, prunedtree_insample_preds)
insample_prtree_accuracy <- sum(diag(prunedtree_conf_matrix_insample)) / sum(prunedtree_conf_matrix_insample)
insample_prtree_accuracy


### random forest ###
install.packages("randomForest")
library(randomForest)
rf_model <- randomForest(x = training_bookingsdf[,-15], y = training_bookingsdf$booking.status, ntree = 100)
rf_predictions <- predict(rf_model, newdata = testing_bookingsdf[,-15])

rf_conf_matrix <- table(testing_bookingsdf$booking.status,rf_predictions)
rf_conf_matrix
rf_accuracy <- sum(diag(rf_conf_matrix)) / sum(rf_conf_matrix)
rf_accuracy
var_importance <- importance(rf_model)
var_importance_sorted <- var_importance[order(var_importance, decreasing = TRUE), ]

print(var_importance_sorted)

rf_model_reduced <- randomForest(x = training_bookingsdf[,c(-2,-6,-10,-11,-12,-15)], y = training_bookingsdf$booking.status, ntree = 50)
rf_predictions_new <- predict(rf_model_reduced, newdata = testing_bookingsdf[,c(-2,-6,-10,-11,-12,-15)])

rf_conf_matrix_new <- table(testing_bookingsdf$booking.status,rf_predictions_new)
rf_reduced_accuracy <- sum(diag(rf_conf_matrix_new)) / sum(rf_conf_matrix_new)
rf_reduced_accuracy

var_importance_final <- importance(rf_model_reduced)
fvar_importance_sort<- var_importance_final[order(var_importance_final, decreasing = TRUE), ]

print(fvar_importance_sort)
par(mfrow =c(1,1))


library(caret)

# Evaluating Random forest with cross-validation on the full dataset
rf_model_cv <- train(
  x = bookings_df[, -15], 
  y = bookings_df$booking.status,
  method = "rf",
  trControl = trainControl(method = "cv", number = 5))

print(paste("CV Random Forest Accuracy:", mean(rf_model_cv$results$Accuracy)))


barplot(var_importance_sorted, horiz = FALSE, las = 2, main = "Random Forest Variable Importance", cex.names = 0.65, ylim = c(0,200), col = "lightblue")


logistic_roc <- roc(response=testing_bookingsdf$booking.status, predictor = logistic_preds)
plot(logistic_roc, legacy.axes = TRUE)




##### Clustering  #####
bookingsscaled_clust <- bookings_df[,-15]
bookingsscaled_clust$number.of.adults <- scale(bookingsscaled_clust$number.of.adults)
bookingsscaled_clust$number.of.children <- scale(bookingsscaled_clust$number.of.children)
bookingsscaled_clust$number.of.weekend.nights <- scale(bookingsscaled_clust$number.of.weekend.nights)
bookingsscaled_clust$number.of.week.nights <- scale(bookingsscaled_clust$number.of.week.nights)
bookingsscaled_clust$lead.time <- scale(bookingsscaled_clust$lead.time)
bookingsscaled_clust$P.C <- scale(bookingsscaled_clust$P.C)
bookingsscaled_clust$P.not.C <- scale(bookingsscaled_clust$P.not.C )
bookingsscaled_clust$average.price <- scale(bookingsscaled_clust$average.price)
bookingsscaled_clust$special.requests <- scale(bookingsscaled_clust$special.requests)

normalized_data <- apply(bookingsscaled_clust[,c(1,2,3,4,8,11,12,13,14)], 2, function(x) (x - min(x)) / (max(x) - min(x)))

install.packages("factoextra")
library(factoextra)
ward_elbow <- fviz_nbclust(bookings_df[,-15], hcut, method = "wss", k.max = 8)
print(ward_elbow)

ward_elbow2 <- fviz_nbclust(normalized_data, hcut, method = "wss", k.max = 8)
print(ward_elbow2)


library(cluster)
gower_dist <- daisy(bookings_df[,-15], metric = "gower")
gower_dist2 <- daisy(normalized_data, metric = "gower")


### ward method
hcluster_ward <- hclust(gower_dist, method = "ward.D")
plot(hcluster_ward)
rect.hclust(hcluster_ward, k=2, border="red")
rect.hclust(hcluster_ward, k=3, border="green")

ward_2clusters <- cutree(hcluster_ward, k = 2)

sil_ward_2clusters <- silhouette(ward_2clusters, gower_dist)

par(mfrow =c(2,1))
fviz_silhouette(sil_ward_2clusters, main = "Silhouette Plot for Hierarchical Clustering with Ward's Method(Uncaled Data)")
sum(sil_ward_2clusters < 0)
cluster.stats(gower_dist, ward_2clusters)$dunn


### ward method with normalized data
hcluster_ward2 <- hclust(gower_dist2, method = "ward.D")
plot(hcluster_ward2)
rect.hclust(hcluster_ward2, k=2, border="red")
rect.hclust(hcluster_ward2, k=3, border="green")

ward_2clusters_scaled <- cutree(hcluster_ward2, k = 2)

sil_ward_2clusters_sc <- silhouette(ward_2clusters_scaled, gower_dist2)

par(mfrow =c(1,2))
fviz_silhouette(sil_ward_2clusters, main = "Silhouette Plot for Hierarchical Clustering with Ward's Method(Uncaled Data)")
fviz_silhouette(sil_ward_2clusters_sc, main = "Silhouette Plot for Hierarchical Clustering with Ward's Method (Scaled Data)")
sum(sil_ward_2clusters_sc < 0)

cluster.stats(gower_dist2, ward_2clusters_scaled)$dunn

### complete method
par(mfrow =c(1,1))
hcluster_complete <- hclust(gower_dist, method = "complete")
plot(hcluster_complete)
rect.hclust(hcluster_complete, k=2, border="red")
rect.hclust(hcluster_complete, k=3, border="green")

complete_2clusters <- cutree(hcluster_complete, k = 2)

sil_comp_2clusters <- silhouette(complete_2clusters, gower_dist)

fviz_silhouette(sil_comp_2clusters, main = "Silhouette Plot for Hierarchical Clustering with Complete Method")
sum(sil_comp_2clusters < 0)

cluster.stats(gower_dist, complete_2clusters)$dunn

### complete method with normalized data
hcluster_complete2 <- hclust(gower_dist2, method = "complete")
plot(hcluster_complete2)
rect.hclust(hcluster_complete2, k=2, border="red")
rect.hclust(hcluster_complete2, k=3, border="green")

complete_2clusters_scaled <- cutree(hcluster_complete2, k = 2)

sil_comp_2clusters_sc <- silhouette(complete_2clusters_scaled, gower_dist2)


fviz_silhouette(sil_comp_2clusters_sc, main = "Silhouette Plot for Hierarchical Clustering with Complete Method (scaled data, 2 clusters)")
sum(sil_comp_2clusters_sc < 0)

cluster.stats(gower_dist2, complete_2clusters_scaled)$dunn


complete_3clusters_scaled <- cutree(hcluster_complete2, k = 3)
sil_comp_3clusters_sc <- silhouette(complete_3clusters_scaled, gower_dist2)
fviz_silhouette(sil_comp_3clusters_sc, main = "Silhouette Plot for Hierarchical Clustering with Complete Method (scaled data, 3 clusters)")
sum(sil_comp_3clusters_sc < 0)



### average method##
hcluster_average <- hclust(gower_dist, method = "average")
plot(hcluster_average)
rect.hclust(hcluster_average, k=2, border="red")
rect.hclust(hcluster_average, k=3, border="red")

cluster_assignments_av <- cutree(hcluster_average, k = 2)

sil_widths3 <- silhouette(cluster_assignments_av, gower_dist)

fviz_silhouette(sil_widths3, main = "Silhouette Plot for Hierarchical Clustering with Average Method")
sum(sil_widths3 < 0)


### elbow method for cluster selection for k-medoids clustering ###
kmed_indices <- fviz_nbclust(normalized_data, pam, method = "wss", k.max = 8)
print(kmed_indices)

#### K-Medoids ####
kmedoids_2clusters <- pam(normalized_data, 2)
kmed_2clusters_assign <- kmedoids_2clusters$clustering

sil_kmed_2clusters <- silhouette(kmed_2clusters_assign, gower_dist2)
fviz_silhouette(sil_kmed_2clusters, main = "Silhouette Plot for k-medoids(2 clusters)")
negative_silvalues_number <- sum(sil_kmed_2clusters < 0)
print(negative_silvalues_number)

cluster.stats(gower_dist2, kmed_2clusters_assign)$dunn


kmedoids_4clusters <- pam(normalized_data, 4)
kmed_4clusters_assign <- kmedoids_4clusters$clustering

sil_kmed_4clusters <- silhouette(kmed_4clusters_assign, gower_dist2)
fviz_silhouette(sil_kmed_4clusters, main = "Silhouette Plot for k-medoids(4 clusters)")
negative_silvalues_number4 <- sum(sil_kmed_4clusters < 0)
print(negative_silvalues_number4)



boxplot(bookings_df$number.of.adults ~ complete_2clusters, col = c("blue", "red"))  ## significant difference
boxplot(bookings_df$number.of.children ~ complete_2clusters, col = c("blue", "red"))
boxplot(bookings_df$month_of_reservation ~ complete_2clusters, col = c("blue", "red")) ## seems insignificant difference I will do anova to be sure
boxplot(bookings_df$lead.time ~ complete_2clusters, col = c("blue", "red")) ## significant difference
boxplot(bookings_df$P.C ~ cluster_assignments_com, col = c("blue", "red")) ## significant difference
boxplot(bookings_df$P.not.C ~ complete_2clusters, col = c("blue", "red")) ## significant difference
boxplot(bookings_df$average.price ~ complete_2clusters, col = c("blue", "red")) ## significant difference
boxplot(bookings_df$special.requests ~ complete_2clusters, col = c("blue", "red")) ## seems insignificant difference I will do anova to be sure
boxplot(bookings_df$number.of.weekend.nights ~ complete_2clusters, col = c("blue", "red"))
boxplot(bookings_df$number.of.week.nights ~ complete_2clusters, col = c("blue", "red"))



aov_requests <- aov(bookings_df$special.requests ~ cluster_assignments_com)
summary(aov_requests) ## p-value = 0.663 > 0.05 I do not reject the equality of the means of
## special requests variable between the 2 clusters

aov_kids <- aov(bookings_df$number.of.children ~ complete_2clusters)
summary(aov_kids)

aov_PC <- aov(bookings_df$P.C ~ complete_2clusters)
summary(aov_PC)

aov_weeknights <- aov(bookings_df$number.of.week.nights ~ complete_2clusters)
summary(aov_weeknights)



gower_dist3 <- daisy(bookings_df[,c(-2,-4,-5,-14,-15,-16)], metric = "gower")

hcluster_wardf <- hclust(gower_dist3, method = "ward.D")
plot(hcluster_wardf)
rect.hclust(hcluster_wardf, k=2, border="red")
rect.hclust(hcluster_wardf, k=3, border="green")

clusters_wardfinal <- cutree(hcluster_wardf, k = 2)

sil_widthsfinal <- silhouette(clusters_wardfinal, gower_dist3)

fviz_silhouette(sil_widthsfinal, main = "Silhouette Plot for Hierarchical Clustering with Ward's Method")
sum(sil_widthsfinal < 0)
cluster.stats(gower_dist3, clusters_wardfinal)$dunn


### Diagrams for interpretation
boxplot(bookings_df$number.of.adults ~ clusters_wardfinal, col = c("blue", "red"), lab = "Number of Adults", xlab = "Clusters", main="Adults per Cluster")  ## significant difference
boxplot(bookings_df$lead.time ~ clusters_wardfinal, col = c("blue", "red"), ylab = "Lead Time", xlab = "Clusters", main="Lead Time per Cluster") ## significant difference
boxplot(bookings_df$average.price ~ clusters_wardfinal, col = c("blue", "red"), ylab = "Average Price", xlab = "Clusters", main="Average Price per Cluster") ## significant difference
boxplot(bookings_df$number.of.weekend.nights ~ clusters_wardfinal, col = c("blue", "red"), ylab = "Number of weekend nights", xlab = "Clusters", main="Weekend Nights per Cluster") ## significant difference
boxplot(bookings_df$number.of.week.nights ~ clusters_wardfinal, col = c("blue", "red"), ylab = "Number of week nights", xlab = "Clusters", main="Week Nights per Cluster") ## insignificant difference I will try to remove it

table(bookings_df$car.parking.space, clusters_wardfinal)
table(bookings_df$room.type, clusters_wardfinal)
table(bookings_df$repeated, clusters_wardfinal)
table(bookings_df$market.segment.type, clusters_wardfinal)
