Sub Main
	Call JoinDatabase()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Archivo: Unir bases de datos
Function JoinDatabase
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Ejemplox-Clientes.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "NUM_CLI", "NUM_CLI", "A"
	task.CreateVirtualDatabase = False
	dbName = "Tabla_unida01.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function