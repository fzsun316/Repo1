#use 'xlrd' package to read data from Excel files
#to get 'xlrd' package, download from: https://pypi.python.org/pypi/xlrd
from xlrd import open_workbook
#use 'xlwt' package to write data to Excel files
#to get 'xlwt' package, download from: https://pypi.python.org/pypi/xlwt
import xlwt
import math
from operator import itemgetter
#get random rows to create random data sets, to use Bagging Method
import random

#main() function: read from train data set to create decision trees.
#read from evaluation data set to get the outputs
def main():
	#train data set. To change source file, please change the file name here!
	wb_train=open_workbook('train.xlsx','rb')
	sheet_train=wb_train.sheets()[0]
	
	NUM_ROWS=300  #number of rows to create a single decision tree
	DECISION_TREE_NUM=1  #number of decision trees to create
	COLUMN_DATA_BEGIN=0  #the column that data begins, avoid the ID columns
	
	dataset_train=[]  #list to get the training data
	random_row=[]  
	for i in range(DECISION_TREE_NUM): #get random rows of training data for each tree
		random_row_single=[]
		for j in range(NUM_ROWS):
			print 'NUM_ROWS1 ',j
			random_row_single.append(random.randint(0,sheet_train.nrows-1))
		random_row.append(random_row_single)
	
	decision_tree=[]
	for i in range(DECISION_TREE_NUM): #get training data for each tree
		dataset_train_single=[]
		for j in range(NUM_ROWS):
			print 'NUM_ROWS2 ',j
			sLine=[]
			for col in range(COLUMN_DATA_BEGIN, sheet_train.ncols):
				sLine.append(sheet_train.cell(random_row[i][j],col).value)
			dataset_train_single.append(sLine)
		decision_tree.append(learn(dataset_train_single))
	
	
	
	#evaluation data set. To change source file, please change the file name here!
	wb_evaluation=open_workbook('test.xlsx','rb')
	sheet_evaluation=wb_evaluation.sheets()[0]
	
	dataset_evaluation=[]
	for row in range(sheet_evaluation.nrows): #get evaluation data
		print 'sheet_evaluation.nrows ',row
		sLine=[]
		for col in range(COLUMN_DATA_BEGIN, sheet_evaluation.ncols):
			sLine.append(sheet_evaluation.cell(row,col).value)
		dataset_evaluation.append(sLine)
	
	answer=[]  #output list for evaluation data set
	for row in dataset_evaluation:
		print 'dataset_evaluation ',row
		answer_tmp=[]
		for tree in range(DECISION_TREE_NUM):  #use every decision tree to evaluate the data
			answer_tmp.append(classify(row,decision_tree[tree]))
		#answer_tmp.append(classify(row,decision_tree[0]))
		answer.append(majortiyCnt(answer_tmp))  #Bagging method. Voting for the output results
	

	#write the output to a specified file('output.xls')
	wb_output=xlwt.Workbook()
	sheet_output = wb_output.add_sheet('Output')
	for i in range(len(answer)):
		sheet_output.write(i,0,answer[i])
	wb_output.save('output2.xls')
	print 'decision_tree: ',decision_tree
	
	

def learn(dataset,depth=0):
	classList = [value[-1] for value in dataset]
	if depth>160:  #max depth of tree: number of features/2
		return majortiyCnt(classList)
	if classList.count(classList[0])==len(classList):  #all classes remaining are the same
		return classList[0]
	
	best_tmp = choose_best_feature(dataset)  #ID3 algorithm: get the attribute with most info gain 
	best_feature = best_tmp[0]  #best attribute with most info gain
	best_value = best_tmp[1]  #best split of the best attribute
	
	decision_tree = {best_feature:{}}  #use Dictionary to build decision tree
	sub_dataset=split_dataset(dataset,best_feature,best_value) #split the data set into subsets using the attribute
	
	depth+=1
	#build decision tree for the data with less value of the attribute
	if len(sub_dataset[0])!=0:
		decision_tree[best_feature][best_value] = learn(sub_dataset[0],depth)  #recursion
	#build decision tree for the data with bigger value of the attribute
	if len(sub_dataset[1])!=0:
		decision_tree[best_feature][best_value+1] = learn(sub_dataset[1],depth)  #recursion
	return decision_tree

def classify(dataset, decision_tree):  #classify the dataset with specific decision tree
	if isinstance(decision_tree, (int,float)):  #only class value left: classify ends
		return decision_tree
	best_feature = decision_tree.keys()[0] #get the split attribute from tree
	best_value = decision_tree.values()[0] #get the split value from tree
	
	del dataset[best_feature]
	if dataset[best_feature]<=best_value.keys()[0]: #value less than the split value
		return classify(dataset,best_value.values()[0])
	else: #value bigger than the split value
		return classify(dataset,best_value.values()[1])



def calc_entropy(dataset):  #calculate the entropy of dataset
	lines=len(dataset);
	labels={}
	for row in dataset:
		labels.setdefault(row[-1],0)
		labels[row[-1]]+=1
	entropy=0.0
	for key in labels.keys():
		prob=float(labels[key])/lines
		entropy-=math.log(prob,2)*prob
	return entropy
	
def split_dataset(dataset, cIndex, cValue): #choose the data set, that the value of attribute cIndex is cValue
	small_set=[]
	big_set=[]
	for row in dataset:
		tmp_set=row[:cIndex]
		tmp_set.extend(row[cIndex+1:])
		if row[cIndex]<=cValue:
			small_set.append(tmp_set)
		else:
			big_set.append(tmp_set)
	return [small_set,big_set]
	
def choose_best_feature(dataset): #ID3 algorithm: get the attribute with most info gain 
	base_entropy = calc_entropy(dataset) #calculate the entropy of class values
	best_feature=-1
	best_value=0
	best_info_gain=-9999.0
	for f in range(len(dataset[0])-1): #calculate best info gain of every attribute
		features=[value[f] for value in dataset]
		features_set = set(features)
		features_set=list(features_set)
		new_entropy = 9999.0
		best_value_tmp=0
		for v in features_set[:len(features_set)-1]: #deal with continuous value, try every split, to get the best split value
			small_set = split_dataset(dataset,f,v)[0]
			small_set_prob=len(small_set)/float(len(dataset))
			big_set = split_dataset(dataset,f,v)[1]
			big_set_prob=len(big_set)/float(len(dataset))
			new_entropy_tmp = small_set_prob*calc_entropy(small_set)+big_set_prob*calc_entropy(big_set)
			
			if new_entropy>new_entropy_tmp:  #get the split value with least entropy value
				new_entropy=new_entropy_tmp
				best_value_tmp=v
			
		info_gain_tmp =	base_entropy-new_entropy
		if info_gain_tmp>best_info_gain:
			best_info_gain=info_gain_tmp
			best_feature=f
			best_value=best_value_tmp
	return [best_feature,best_value]
	
def majortiyCnt(classList):  #Bagging: vote for the most class value
	classCount={}
	for vote in classList:
		classCount.setdefault(vote,0)
		classCount[vote]+=1
	sortedclassCount = sorted(classCount.iteritems(),key=itemgetter(1),reverse=True)
	return sortedclassCount[0][0]
	
		
if __name__=='__main__':  #call the main() function
	main()

