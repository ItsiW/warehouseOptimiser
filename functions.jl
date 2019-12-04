function readData(inputFileName)
    capacities = readxlsheet(inputFileName, "Containers", skipstartrows=1)[:,2]
    containers = collect(1:length(capacities))
    containernames = readxlsheet(inputFileName, "Containers", skipstartrows=1)[:,1]
    itemcodes = readxlsheet(inputFileName, "Products", skipstartrows=1)[:,1]
    itemdescs = readxlsheet(inputFileName, "Products", skipstartrows=1)[:,2]
    products = collect(1:length(itemdescs))
    prodDict = Dict([(itemdescs[i] => i) for i in products])
    marknames = readxlsheet(inputFileName, "P4 Orders", skipstartrows=1)[:,1]
    marknames = unique(marknames)
    markets = collect(1:length(marknames))
    markDict = Dict([(marknames[i] => i) for i in markets])
    StartDate = Date(readxlsheet(inputFileName, "Dates", skipstartrows=1)[1])
    demandlist = readxlsheet(inputFileName, "P4 Orders", skipstartrows=1)
    # create demand matrix
    demand = zeros(length(products),length(markets))
    for d in 1:length(demandlist[:,1])
        demand[get(prodDict, demandlist[d,3],1), get(markDict, demandlist[d,1],1)] = demandlist[d,4]
    end
    volumes = readxlsheet(inputFileName, "Products", skipstartrows=1)[:,3]
    totvolume = [sum(volumes[i]*demand[i,j] for i in products) for j in markets]
    # calculate allocation of containers
    allocation = zeros(length(markets),length(containers))
    for j in markets
        allocation[j,1] = floor.(totvolume[j]/capacities[1])
        remaining = totvolume[j] - allocation[j,1]*capacities[1]
        while remaining >0
            next = 1
            for k in containers
                if remaining <= capacities[k]
                    next = k
                else
                    break
                end
            end
            allocation[j,next] += 1
            remaining -= capacities[next]
        end
    end
    values = readxlsheet(inputFileName, "Products", skipstartrows=1)[:,4]
    clearing = readxlsheet(inputFileName, "Products", skipstartrows=1)[:,5]

    # process non-working days
    packingdays = readxlsheet(inputFileName, "Dates", skipstartrows=1)[:,2]
    days = collect(1:length(packingdays))
    weekend = Int64[]
    for t in days
        if packingdays[t] != "Yes"
            weekend = [weekend;t]
        end
    end

    # create supply-date matrix
    supplylist = [readxlsheet(inputFileName, "Schedule", skipstartrows=1);readxlsheet(inputFileName, "Stock On Hand", skipstartrows=1)]
    supply = zeros(length(products),28)
    for i in 1:length(supplylist[:,1])
        supplylist[i,3] = Dates.days(Date(supplylist[i,3]) - StartDate) + 1 + convert(Int64, clearing[get(prodDict, supplylist[i,2],1)])
    end
    for s in 1:size(supplylist)[1]
        if supplylist[s,3] > length(days) + 0.1
        elseif supplylist[s,3]>1.1
            supply[get(prodDict, supplylist[s,2],1), supplylist[s,3]] += supplylist[s,4]
        else
            supply[get(prodDict, supplylist[s,2],1), 1] += supplylist[s,4]
        end
    end
    
    # check supply meets demand
    errors = 0
    errorlist = String[]
    for i in products
        if (round(sum(demand[i,j] for j in markets)) > round(sum(supply[i,t] for t in days)))
            errors += 1
            errorlist = [errorlist;itemdescs[i]]
        end
    end
    if errors > 0.1
        open(outputFileName2, "w") do f
            write(f, "ERROR Demand has not been met for the following products:")
            for i in 1:errors
                write(f, "\n$(errorlist[i])")
            end
        end
        quit()
    end

    products, markets, days, containers, demand, supply, allocation, capacities, volumes, values, marknames, itemcodes, itemdescs, StartDate, containernames, weekend, totvolume
end

function writeDemandSupply(products, markets, days, containers, demand, supply, allocation, capacities, volumes, values, 
marknames, itemcodes, itemdescs, StartDate, containernames, weekend, totvolume)
    open(outputFileName1, "w") do f
        write(f, "DEMAND\n")
        write(f, ",,Volume,Value")
        for j in markets
            write(f, ",$(marknames[j])")
        end
        write(f, ",Item Units,Item Volume,Item Value")

        for i in products
            write(f, "\n$(itemcodes[i]),$(itemdescs[i]),$(volumes[i]),\$$(round(values[i],digits=2))")
            for j in markets
                write(f, ",$(demand[i,j])")
            end
            write(f, ",$(sum(demand[i,j] for j in markets)),$(round(sum(demand[i,j]*volumes[i] for j in markets),digits=3)),\$$(round(sum(demand[i,j]*values[i] for j in markets),digits=2))")
        end
        write(f, "\n,,,Customer Units")
        for j in markets
            write(f, ",$(sum(demand[i,j] for i in products))")
        end
        write(f, "\n,,,Customer Volume")
        for j in markets
            write(f, ",$(round(totvolume[j],digits=3))")
        end
        write(f, "\n,,,Customer Value")
        for j in markets
            write(f, ",\$$(round(sum(values[i]*demand[i,j] for i in products),digits=2))")
        end
        for k in containers
            write(f, "\n,,,$(containernames[k])")
            for j in markets
                write(f, ",$(allocation[j,k])")
            end
        end

        write(f, "\n\nREADY TO PACK")
        write(f, "\n,,Volume,Value")
        for t in days
            write(f, ",$(Dates.format(StartDate + Dates.Day(t-1), "d u"))")
        end
        write(f, ",Item Units,Item Volume,Item Value")
        for i in products
            write(f, "\n$(itemcodes[i]),$(itemdescs[i]),$(volumes[i]),\$$(round(values[i],digits=2))")
            for t in days
                write(f, ",$(supply[i,t])")
            end
            write(f, ",$(sum(supply[i,t] for t in days)),$(round(sum(supply[i,t]*volumes[i] for t in days),digits=3)),\$$(round(sum(supply[i,t]*values[i] for t in days),digits=2))")
        end
        write(f, "\n,,,Supply Units")
        for t in days
            write(f, ",$(round(sum(supply[i,t] for i in products)))")
        end
        write(f, "\n,,,Supply Volume")
        for t in days
            write(f, ",$(round(sum(volumes[i]*supply[i,t] for i in products),digits=3))")
        end
        write(f, "\n,,,Supply Value")
        for t in days
            write(f, ",\$$(round(sum(values[i]*supply[i,t] for i in products),digits=2))")
        end
    end
end

function writeSchedule(products, markets, days, containers, demand, supply, allocation, capacities, volumes, values, 
marknames, itemcodes, itemdescs, StartDate, containernames, weekend, xmatrix, ymatrix, zmatrix)  
    markdemand = zeros(length(markets))
    markvolume = zeros(length(markets))
    markvalue = zeros(length(markets))
    for j in markets
        markdemand[j] = sum(demand[i,j] for i in products)
        markvolume[j] = sum(demand[i,j]*volumes[i] for i in products)
        markvalue[j] = sum(demand[i,j]*values[i] for i in products)
    end

    daydemand = zeros(length(days))
    dayvolume = zeros(length(days))
    dayvalue = zeros(length(days))
    initdemand = sum(demand[i,j] for i in products for j in markets)
    initvolume = sum(demand[i,j]*volumes[i] for i in products for j in markets)
    initvalue = sum(demand[i,j]*values[i] for i in products for j in markets)

    open(outputFileName2, "w") do f
        write(f, "Initial Demand,")
        for j in markets
            write(f, ",$(marknames[j])")
        end
        write(f, ",Total\n,Units")
        for j in markets
            write(f, ",$(round(markdemand[j]))")
        end
        write(f, ",$(round(initdemand))\n,Volume")
        for j in markets
            write(f, ",$(round(markvolume[j],digits=3))")
        end
        write(f, ",$(round(initvolume,digits=3))\n,Value")
        for j in markets
            write(f, ",\$$(round(markvalue[j],digits=2))")
        end
        write(f, ",\$$(round(initvalue,digits=2))\n\n\n")

        for t in days
            pack = 0
            for j in markets
                for k in containers
                    if xmatrix[j,t,k]>0.9
                        pack = 1
                        break
                    end
                end
                if pack == 1
                    break
                end
            end
            if pack <0.1
                write(f, "\n$(Dates.format(StartDate + Dates.Day(t-1), "d u")),No packing\n")
                if t == 1
                    daydemand[t] = initdemand
                    dayvolume[t] = initvolume
                    dayvalue[t] = initvalue
                else
                    daydemand[t] = daydemand[t-1]
                    dayvolume[t] = dayvolume[t-1]
                    dayvalue[t] = dayvalue[t-1]
                end
            else
                write(f, "\n$(Dates.format(StartDate + Dates.Day(t-1), "d u")),")
                for k in containers
                    write(f, ",$(containernames[k])")
                end
                write(f, ",")

                dayproducts = Int64[]
                for i in products
                    for j in markets
                        if zmatrix[i,j,t] > 0.9
                            dayproducts = [dayproducts; i]
                            write(f, ",$(itemdescs[i])")
                            break
                        end
                    end
                end
                write(f, ",,Units,Volume,Value\n")

                for j in markets
                    count = 0
                    for k in containers
                        if xmatrix[j,t,k] > 0.9
                            count = 1
                            break
                        end
                    end
                    if count > 0.9
                        write(f, ",$(marknames[j])")
                        for k in containers
                            write(f, ",$(round(xmatrix[j,t,k]))")
                        end
                        write(f, ",")
                        for i in dayproducts
                            write(f, ",$(zmatrix[i,j,t])")
                            markdemand[j] -= zmatrix[i,j,t]
                            markvolume[j] -= zmatrix[i,j,t]*volumes[i]
                            markvalue[j] -= zmatrix[i,j,t]*values[i]
                        end
                        write(f, ",,$(round(sum(zmatrix[i,j,t] for i in dayproducts))),$(round(sum(zmatrix[i,j,t]*volumes[i] for i in dayproducts),digits=3)),\$$(round(sum(zmatrix[i,j,t]*values[i] for i in dayproducts),digits=2))\n")
                    end
                end
                for k in containers
                    write(f, ",")
                end
                write(f, ",,Units")
                for i in dayproducts
                    write(f, ",$(round(sum(zmatrix[i,j,t] for j in markets)))")
                end
                write(f, "\n")
                for k in containers
                    write(f, ",")
                end
                write(f, ",,Volume")
                for i in dayproducts
                    write(f, ",$(round(sum(zmatrix[i,j,t]*volumes[i] for j in markets),digits=3))")
                end
                write(f, "\n")
                for k in containers
                    write(f, ",")
                end
                write(f, ",,Value")
                for i in dayproducts
                    write(f, ",\$$(round(sum(zmatrix[i,j,t]*values[i] for j in markets),digits=2))")
                end

                for j in markets
                    daydemand[t] += markdemand[j]
                    dayvolume[t] += markvolume[j]
                    dayvalue[t] += markvalue[j]
                end
                write(f, "\n\n,Remaining")
                for j in markets
                    write(f, ",$(marknames[j])")
                end
                write(f, ",Total\n,Units")
                for j in markets
                    write(f, ",$(round(markdemand[j]))")
                end
                write(f, ",$(round(daydemand[t]))\n,Volume")
                for j in markets
                    write(f, ",$(round(markvolume[j],digits=3))")
                end
                write(f, ",$(round(dayvolume[t],digits=3))\n,Value")
                for j in markets
                    write(f, ",\$$(round(markvalue[j],digits=2))")
                end
                write(f, ",\$$(round(dayvalue[t],digits=2))\n\n\n")
            end
        end
    end
    daydemand,dayvolume,dayvalue,initdemand,initvolume,initvalue
end

function writeSummary(products, markets, days, containers, demand, supply, allocation, capacities, volumes,
        values, marknames, itemcodes, itemdescs, StartDate, containernames, weekend, daydemand, dayvolume,
        dayvalue, initdemand, initvolume, initvalue)    
    open(outputFileName3, "w") do f
        write(f, "Remaining at the end of day\n")
        for k in containers
            write(f, ",$(containernames[k])")
        end
        write(f, ",,Units,Units%,,Volume,Volume%,,Value,Value%\nTotal,,,,$(round(initdemand)),100%,,$(round(initvolume,digits=3)),100%,,\$$(round(initvalue,digits=2)),100%")
        for t in days
            write(f, "\n$(Dates.format(StartDate + Dates.Day(t-1), "d u"))")
            for k in containers
                daycontainers = 0
                for j in markets
                    daycontainers += xmatrix[j,t,k]
                end
                write(f, ",$(round(daycontainers))")
            end
            write(f, ",,$(round(daydemand[t])),$(round(daydemand[t]*100/initdemand,digits=2))%,,$(round(dayvolume[t],digits=3)),$(round(dayvolume[t]*100/initvolume,digits=2))%,,\$$(round(dayvalue[t],digits=2)),$(round(dayvalue[t]*100/initvalue,digits=2))%")
        end
    end
end
