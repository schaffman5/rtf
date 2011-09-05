require(rtf)
require(lattice)

output<-"test.rtf.doc"
png.res<-300

rtf<-RTF(output,width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
addHeader(rtf,title="Test",subtitle="2011-08-15\n")
addPlot(rtf,plot.fun=plot,width=6,height=6,res=300, iris[,1],iris[,2])

# Try trellis plots

# single page trellis objects
addPageBreak(rtf, width=11,height=8.5,omi=c(0.5,0.5,0.5,0.5))

p <- histogram( ~ height | voice.part, data = singer, xlab="Height")
addTrellisObject(rtf,trellis.object=p,width=10,height=7.5,res=png.res)

p <- densityplot( ~ height | voice.part, data = singer, xlab = "Height")
addTrellisObject(rtf,trellis.object=p,width=10,height=7.5,res=png.res)

# multipage trellis object
p2<-xyplot(uptake ~ conc | Plant, CO2, layout = c(2,2))
addTrellisObject(rtf,trellis.object=p2,width=6,height=6,res=png.res)

addPageBreak(rtf, width=6,height=10,omi=c(0.5,0.5,0.5,0.5))
addTable(rtf,as.data.frame(head(iris)),font.size=9,row.names=FALSE,NA.string="-",col.widths=rep(1,5))

tab<-table(iris$Species,floor(iris$Sepal.Length))
names(dimnames(tab))<-c("Species","Sepal Length")
addParagraph(rtf,"\n\nHere's a new paragraph with another table:\n")
addTable(rtf,tab,font.size=10,row.names=TRUE,NA.string="-",col.widths=c(1,rep(0.5,4)) )

done(rtf)

# Open in word processor
#system(paste("open",output))
