# -*- coding: utf-8 -*-
"""
Created on Fri Apr 26 17:27:35 2019

@author: Yiping Liu
"""
def optimize(inputfile,outputfile):
    
    # import packages
    import pandas as pd
    from gurobipy import Model,GRB
    
    # read input data
    df_norm = pd.read_excel(inputfile,sheet_name='Product',index_col = 0)
    df2 = pd.read_excel(inputfile,sheet_name='Product_Category',index_col = 0)
    df3 = pd.read_excel(inputfile,sheet_name='Product_Subcategory',index_col = 0)
    score_weight = pd.read_excel(inputfile,sheet_name='score_weight',index_col = 0)
    category_req = pd.read_excel(inputfile,sheet_name='category_req',index_col = 0)
    subcategory_req = pd.read_excel(inputfile,sheet_name='subcategory_req',index_col = 0)
    inno_req = pd.read_excel(inputfile,sheet_name='inno_req')
    
    # calculate product quality score based on the weight assigned by user
    df_norm['score'] = score_weight.loc['sales (base)','weight']*df_norm["sales"]\
                          -score_weight.loc['returns (-)','weight']*df_norm['returns']\
                          -score_weight.loc['distribution cost (-)','weight']*df_norm['total_distribution_cost']\
                          +score_weight.loc['margin (+)','weight']*df_norm['margin']\
                          -score_weight.loc['manufacturing capacity (-)','weight']*df_norm['pc0.95']
    
    # gurobi optimization
    mod = Model()

    I = df_norm.index
    J = df2.columns
    K = df3.columns
    
    # decision variable: whether to select product i
    x = mod.addVars(I,vtype = GRB.BINARY)

    # maximize product quality score
    mod.setObjective(sum(x[i]*df_norm.loc[i,'score'] for i in I),\
                     sense = GRB.MAXIMIZE)
    
    # small format quantity constraint
    mod.addConstr(sum(x[i] for i in I) <= 250)
    
    # innovation constraint
    # include at least #(decided by user) of 2018-innovation products
    mod.addConstr(sum(df_norm.loc[i,'innovation_2018']*x[i] for i in I) >= inno_req.loc[0,'Innovation Requirement'])
    
    # category constraint
    # include at least #(decided by user) of products for each category
    for j in J:
        mod.addConstr(sum(df2.loc[i,j]*x[i] for i in I) >= category_req.loc[j,'Minimum Requirement'])

    # subcategory constraint
    # include at least #(decided by user) of products for each subcategory
    for k in K:
        mod.addConstr(sum(df3.loc[i,k]*x[i] for i in I) >= subcategory_req.loc[k,'Minimum Requirement'])

    mod.optimize()
    
    # generate a list of selected product BDC for output
    SKU = []
    for i in I:
        if x[i].x != 0:
            SKU.append(i)
            
    choice = pd.DataFrame(SKU,columns=['BDC'])
    
    choice.to_excel(outputfile,index=False)


if __name__=='__main__':
    import sys, os
    if len(sys.argv)!=3:
        print('Correct syntax: python books.py inputfile outputfile')
    else:
        inputfile=sys.argv[1]
        outputfile=sys.argv[2]
        if os.path.exists(inputfile):
            optimize(inputfile,outputfile)
            print(f'Successfully optimized. Results in "{outputfile}"')
        else:
            print(f'File "{inputfile}" not found!')