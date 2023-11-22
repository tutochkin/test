let
    

    f=(y)=>(x)=>List.RemoveNulls(Table.Column(x,y)){0}?,

    xls=(file)=>
            [f=(x)=>
            let
                    a = Text.Split(x,"] "){1},
                    b = if Text.At(a,0)="_" then Text.Replace(a,"_","") else a,
                    c = Text.Split(b,"."){0}
            in
                    c,
            f2=(x)=>Text.Split(x,"] "){1},  
      
            from = Excel.Workbook(file){[Name="Sheet1"]}[Data],
            lst = List.FirstN(List.Skip(List.Skip(from[Column1],
                                    (x)=>x<>null)),(x)=>x<>null),
            nms = List.Transform(lst,f2),
            tbl = Table.Skip(from,(x)=>x[Column1]<>"Время"),
            pr = Table.PromoteHeaders(tbl),
            oldnms = List.Transform(List.Skip(Table.ColumnNames(pr)),f),
            newnms = List.Transform(List.Zip({oldnms,nms}),(x)=>Text.Combine(x," ")),
            ren = List.Zip({List.Skip(Table.ColumnNames(pr)),newnms}),
            to = Table.RenameColumns(pr,ren),
// Обрезка начала данных по указанным дате и времени            
            
            firstDate = Text.Split(from[Column1]{1}," "){0}, 
            firstTime = Text.Split(from[Column1]{1}," "){1},
            first = DateTime.From(firstDate&" "&firstTime),
            rez = Table.Skip(to,List.PositionOf(Table.TransformColumnTypes(to,{{"Время", type datetime}})[Время],first))

            ][rez],

    csv=(file)=>
            [f=(x)=>
            let
                    a = Text.Split(x,"] "){1},
                    b = if Text.At(a,0)="_" then Text.Replace(a,"_","") else a,
                    c = Text.Split(b,"."){0}
            in
                    c,
            f2=(x)=>Text.Split(x,"] "){1},  
      
            from = Csv.Document(file,null,";",null,1251),
            lst = List.FirstN(List.Skip(List.Skip(from[Column1],
                                    (x)=>x<>"")),(x)=>x<>""),
            nms = List.Transform(lst,f2),
            tbl = Table.Skip(from,(x)=>x[Column1]<>"Время"),
            pr = Table.PromoteHeaders(tbl),
            oldnms = List.Transform(List.Skip(Table.ColumnNames(pr)),f),
            newnms = List.Transform(List.Zip({oldnms,nms}),(x)=>Text.Combine(x," ")),
            ren = List.Zip({List.Skip(Table.ColumnNames(pr)),newnms}),
            to = Table.RenameColumns(pr,ren),
// Обрезка начала данных по указанным дате и времени

            firstDate = Text.Split(from[Column1]{1}," "){0}, 
            firstTime = Text.Split(from[Column1]{1}," "){1},
            first = DateTime.From(firstDate&" "&firstTime),
            rez = Table.Skip(to,List.PositionOf(Table.TransformColumnTypes(to,{{"Время", type datetime}})[Время],first))

            ][rez],

    g=(file)=> if file{1}=".xls" then xls(file{0}) else csv(file{0}),

    adres = Excel.CurrentWorkbook(){[Name="Параметры"]}[Content]{0}[Путь к исходным данным],
    from = Folder.Files(adres),
    filtr = Table.SelectRows(from,(r)=>List.Contains({".csv",".xls"},r[Extension]) 
                                       and not Text.Contains(r[Name],"~")),
    tr = List.Buffer(List.Transform(List.Zip({filtr[Content],filtr[Extension]}),g)),
    lst = List.Distinct(List.Combine(List.Transform(tr,Table.ColumnNames))),
    cmb=Table.Combine(tr,lst),
    gr = Table.Group(cmb, "Время", List.Transform(List.RemoveItems(lst,
                                                  {"Время"}),
                                                  (x)=>{x,f(x)})),
    tgrd =Table.TransformColumnTypes(gr, List.Transform(Table.ColumnNames(gr), 
                                         (x)=> if x<>"Время" 
                                         then {x, type number} 
                                         else {x, type datetime})),
    
    dwn = Table.FillDown(tgrd,Table.ColumnNames(tgrd)),
    opd = (Table.FillUp(dwn,List.Skip(lst))),


    nmn = Table.ColumnNames(opd),
    newnms = List.Transform(Table.ColumnNames(opd), (x)=>Text.Split(x," "){0}),
    ren = List.Zip({Table.ColumnNames(opd),newnms}),
    to = Table.RenameColumns(opd,ren),
    Sort = Table.Sort(to,{{"Время", Order.Ascending}}),
    rez =Table.TransformColumnTypes(Sort, List.Transform(Table.ColumnNames(Sort), 
                                         (x)=> if x<>"Время" 
                                         then {x, type number} 
                                         else {x, type datetime}))
in
    rez