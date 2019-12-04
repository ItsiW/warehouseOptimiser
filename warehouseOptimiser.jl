using JuMP
using Cbc
using ExcelReaders
using Dates

Limit = 5
inputFileName = "Input.xlsx"
outputFileName1 = "output/Data.csv"
outputFileName2 = "output/Packing.csv"
outputFileName3 = "output/Summary.csv"

include("functions.jl")

products, markets, days, containers, demand, supply, allocation, capacities, volumes, values, 
marknames, itemcodes, itemdescs, StartDate, containernames, weekend, totvolume = readData(inputFileName)

writeDemandSupply(products, markets, days, containers, demand, supply, allocation, capacities, volumes, values, 
marknames, itemcodes, itemdescs, StartDate, containernames, weekend, totvolume)

m = Model(with_optimizer(Cbc.Optimizer))

@variable(m, x[markets, days, containers]>=0, Int)
@variable(m, y[products, days]>=0, Int)
@variable(m, z[products, markets, days]>=0, Int)

@objective(m, Min, sum(volumes[i]*y[i,t] for i in products for t in days))

@constraint(m,  firstDay[i in products], 
                    y[i,1] == supply[i,1] - sum( z[i,j,1] for j in markets))
@constraint(m,  conservation[i in products, t in days; t!=1], 
                    y[i,t] == y[i,t-1] + supply[i,t] - sum( z[i,j,t] for j in markets))
@constraint(m,  demandMet[i in products, j in markets],
                    sum( z[i,j,t] for t in days) == demand[i,j])
@constraint(m,  containerAllocation[j in markets, k in containers],
                    sum( x[j,t,k] for t in days) == allocation[j,k])
@constraint(m,  maxContainers[t in days],
                    sum( x[j,t,k] for j in markets for k in containers) <= Limit)
@constraint(m,  noWeekends[t in weekend],
                    sum( x[j,t,k] for j in markets for k in containers) == 0)
@constraint(m,  productsFit[j in markets, t in days],
                    sum( volumes[i]*z[i,j,t] for i in products) <= sum( capacities[k]*x[j,t,k] for k in containers))

optimize!(m)
termination_status(m)

xmatrix = value.(x)
ymatrix = value.(y)
zmatrix = value.(z)

daydemand,dayvolume,dayvalue,initdemand,initvolume,initvalue = writeSchedule(products, markets, days, containers, demand, supply, allocation, capacities, volumes, values, 
marknames, itemcodes, itemdescs, StartDate, containernames, weekend, xmatrix, ymatrix, zmatrix)  
                                                    
writeSummary(products, markets, days, containers, demand, supply, allocation, capacities, volumes,values, marknames, 
        itemcodes, itemdescs, StartDate, containernames, weekend, daydemand, dayvolume, dayvalue, initdemand, 
        initvolume, initvalue)
