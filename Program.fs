
/// ---------------------------------------------------------------------
/// fsBacnetWrite
///     jose.vu@kaer.com - 25-Aug-2018
///     the program to write a value to multiple Bacnet points (in different devices)
/// ---------------------------------------------------------------------
open System
open Argu
open System.IO.BACnet
open ExcelDataReader

System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

// predefined type
type BacnetObjectOption<'a> = 
    | Success of 'a
    | Failure of string 
let bindResult f x = 
    match x with 
    | Success x' -> f x'
    | Failure er -> None
let (>==) x f = bindResult f x
let (>>=) x f = Option.bind f x

type BACnetPoints =
    {
        DeviceID  : string
        BacnetObj : BacnetObjectOption<BacnetObjectId>
        Name      : string
        IsOnline  : bool
    }     

// for parsing purpose
type CLIArguments =
    | BacnetIP of string
    | Value of string
    | Filepath of string
    | WritePriority of int      // write priority (default is 8)
    | Timeout of float          // timeout in bacnet whois
with
    interface IArgParserTemplate with
        member s.Usage =
            match s with
            | BacnetIP _ -> "ip for bacnet ip"
            | Value _ -> "value to write, blank is null"
            | Filepath _ -> "Excel file"
            | WritePriority _ -> "BACnet write priority, default is 8"
            | Timeout _ -> "Timeout in seconds, default is 2s"
let parser = ArgumentParser.Create<CLIArguments>(programName = "fsBacnetWrite.exe")

// BACNET
let mutable devList = Map.empty<string,BacnetAddress>
let handlerOnIam (sender:BacnetClient) (adr:BacnetAddress) deviceId maxApdu (segmentation:BacnetSegmentations) vendorId =
    match devList.TryFind (deviceId.ToString()) with
    | None ->
        printfn "[handlerOnIam]: adding %d to devList" deviceId    
        //addDevToList adr deviceId
        devList <- devList.Add(deviceId.ToString(), adr) 
    | _ -> ()
    ()

let default_bacnet_ip = ""

[<EntryPoint>]
let main argv = 
    // Parse the arguments
    let argList = parser.Parse(inputs=argv, ignoreUnrecognized=true)    
    let bacnetip = 
        match argList.GetResult (BacnetIP, default_bacnet_ip) with
        | "" -> failwith "Please ENTER bacnet ip as --bacnetip ipaddress"; ""
        | _ -> argList.GetResult (BacnetIP, default_bacnet_ip)
    let argVal = argList.GetResult (Value, "")    
    let valueToWrite = 
        match argVal with 
        | "" -> BacnetValue( BacnetApplicationTags.BACNET_APPLICATION_TAG_NULL, None )
        | _  -> BacnetValue( BacnetApplicationTags.BACNET_APPLICATION_TAG_REAL, System.Convert.ToSingle argVal )
    let bacnetTimeout = argList.GetResult( Timeout, 2.0 ) * 1000.0
    let bacnetWP = argList.GetResult( WritePriority, 8 )

    // Create bacnet client and send whois
    bacnetip |> printfn "bacnet ip is %s"
    let bacnetClient = new BacnetClient(new BacnetIpUdpProtocolTransport(0xBAC0, false,false, 1476, bacnetip))
    bacnetClient.Start()
    bacnetClient.add_OnIam (new BacnetClient.IamHandler( handlerOnIam ))
    printfn "Send whois..."
    bacnetClient.WhoIs()

    // Read AV/AO points from Excel file
    let filepath = argList.GetResult (Filepath, @"vavAO_list.xlsx")    
    let stream1  = IO.File.Open(filepath,IO.FileMode.Open, IO.FileAccess.Read)
    let reader = ExcelReaderFactory.CreateReader(stream1)
    let result = reader.AsDataSet( new ExcelDataSetConfiguration( ConfigureDataTable = fun (_:IExcelDataReader) -> ExcelDataTableConfiguration( UseHeaderRow = true) ))
    let df = result.Tables.[0] 

    // wait some times for all devices to be discovered
    Async.Sleep( int bacnetTimeout )  |> Async.RunSynchronously //wait 1s

    // read the Excel
    let convertDataRow (x:Data.DataRow) = 
        match devList.ContainsKey(x.["Device-instance"].ToString()) with
        | true -> let objtmp = 
                        match x.["Analog"].ToString().ToUpper() with
                        | "AO" -> new BacnetObjectId(BacnetObjectTypes.OBJECT_ANALOG_OUTPUT, System.Convert.ToUInt32 x.["Object"])
                        | _ -> new BacnetObjectId(BacnetObjectTypes.OBJECT_ANALOG_VALUE, System.Convert.ToUInt32 x.["Object"])
                  { DeviceID = x.["Device-instance"].ToString(); BacnetObj = Success objtmp; Name = sprintf "Dev%s:%s%s" (x.["Device-instance"].ToString()) (x.["Analog"].ToString()) (x.["Object"].ToString()); IsOnline = true  }
        | false -> { DeviceID = x.["Device-instance"].ToString(); BacnetObj = Failure "Device is not on the list"; Name=""; IsOnline=false }                    
    
    let pointlist = df.Rows |> Seq.cast<Data.DataRow> |> Seq.map convertDataRow |> Seq.toList

    let writeFunc (x:BacnetAddress) (y:BacnetObjectId) =
        try
            bacnetClient.WritePropertyRequest( x, y, BacnetPropertyIds.PROP_PRESENT_VALUE, seq{yield valueToWrite} )
        with
        | _ -> false
    let writeToDevice (x:BACnetPoints) = 
        match x.IsOnline with
        | false ->
            printfn "Device %s is not online ...." x.DeviceID; None
        | _ ->
            match argVal with 
            | "" -> printf "Write null to %s" x.Name
            | _ -> printf "Write %s to %s" argVal x.Name
            (devList.TryFind x.DeviceID) >>= (fun bdev -> 
                x.BacnetObj >== (fun bobj ->
                    match (writeFunc bdev bobj) with
                    | true -> printfn " ... OK"; Some "write success"
                    | false -> printfn " ... Fail"; None
                )  
            )    
    bacnetClient.WritePriority <- uint32 bacnetWP
    pointlist |> List.map writeToDevice |> ignore
    0 // return an integer exit code
