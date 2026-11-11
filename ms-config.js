/*
  Conexión del SGI con SharePoint (Microsoft 365 de CWL) para la pestaña Certificaciones.
  Solo hay que completar clientId con el "Id. de aplicación (cliente)" de la app registrada en
  Microsoft Entra ID (portal.azure.com → Microsoft Entra ID → Registros de aplicaciones).
  Estos datos no son secretos: cada usuario entra con su propia cuenta de CWL y ve solo lo que SharePoint le permite.
*/
window.SGI_MS_CONFIG = {
  clientId: "00cded60-0c49-4de1-8c33-be2cff01fa85",               // Id. de aplicación (cliente) de Entra ID
  tenant: "81a055a9-49c5-4266-818f-8e9ccf067e51", // Id. de directorio (inquilino) de CWL
  // Carpeta de certificaciones: SharePoint Operaciones › Pozos › Certificaciones
  site: "cwlms.sharepoint.com:/sites/operaciones",
  library: "Pozos",
  folder: "Certificaciones",
  // Identificadores de esa misma carpeta (aceleran la conexión; si cambian, se usa site/library/folder)
  driveId: "b!kFGiJyDvu02rIJM-tRBJ2J5g2ZI1gEdPp4WUv9kPL1WJbNtI8ufhTZn3kB4Cu5UD",
  folderId: "01VF7QXFIKTA7K77WLEJA36US2TNU3SHG5",
  // Subcarpeta a la que se mueve el certificado anterior cuando se reemplaza
  carpetaVencidas: "certificaciones vencidas",
  // Días de anticipación para marcar "Por vencer"
  diasAviso: 30
};
