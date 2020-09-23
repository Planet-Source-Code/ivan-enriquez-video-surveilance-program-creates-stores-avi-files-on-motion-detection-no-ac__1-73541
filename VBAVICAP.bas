Attribute VB_Name = "AVICAP"

Public Const WM_USER = &H400
Type POINTAPI
        x As Long
        y As Long
End Type
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessageS Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long
Private Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByRef lParam As Any) As Long

Public Const WM_CAP_START = WM_USER

Public Const WM_CAP_GET_CAPSTREAMPTR = WM_CAP_START + 1

Public Const WM_CAP_SET_CALLBACK_ERROR = WM_CAP_START + 2
Public Const WM_CAP_SET_CALLBACK_STATUS = WM_CAP_START + 3
Public Const WM_CAP_SET_CALLBACK_YIELD = WM_CAP_START + 4
Public Const WM_CAP_SET_CALLBACK_FRAME = WM_CAP_START + 5
Public Const WM_CAP_SET_CALLBACK_VIDEOSTREAM = WM_CAP_START + 6
Public Const WM_CAP_SET_CALLBACK_WAVESTREAM = WM_CAP_START + 7
Public Const WM_CAP_GET_USER_DATA = WM_CAP_START + 8
Public Const WM_CAP_SET_USER_DATA = WM_CAP_START + 9
    
Public Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11
Public Const WM_CAP_DRIVER_GET_NAME = WM_CAP_START + 12
Public Const WM_CAP_DRIVER_GET_VERSION = WM_CAP_START + 13
Public Const WM_CAP_DRIVER_GET_CAPS = WM_CAP_START + 14

Public Const WM_CAP_FILE_SET_CAPTURE_FILE = WM_CAP_START + 20
Public Const WM_CAP_FILE_GET_CAPTURE_FILE = WM_CAP_START + 21
Public Const WM_CAP_FILE_ALLOCATE = WM_CAP_START + 22
Public Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23
Public Const WM_CAP_FILE_SET_INFOCHUNK = WM_CAP_START + 24
Public Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25

Public Const WM_CAP_EDIT_COPY = WM_CAP_START + 30

Public Const WM_CAP_SET_AUDIOFORMAT = WM_CAP_START + 35
Public Const WM_CAP_GET_AUDIOFORMAT = WM_CAP_START + 36

Public Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42
Public Const WM_CAP_DLG_VIDEODISPLAY = WM_CAP_START + 43
Public Const WM_CAP_GET_VIDEOFORMAT = WM_CAP_START + 44
Public Const WM_CAP_SET_VIDEOFORMAT = WM_CAP_START + 45
Public Const WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46

Public Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50
Public Const WM_CAP_SET_OVERLAY = WM_CAP_START + 51
Public Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52
Public Const WM_CAP_SET_SCALE = WM_CAP_START + 53
Public Const WM_CAP_GET_STATUS = WM_CAP_START + 54
Public Const WM_CAP_SET_SCROLL = WM_CAP_START + 55

Public Const WM_CAP_GRAB_FRAME = WM_CAP_START + 60
Public Const WM_CAP_GRAB_FRAME_NOSTOP = WM_CAP_START + 61

Public Const WM_CAP_SEQUENCE = WM_CAP_START + 62
Public Const WM_CAP_SEQUENCE_NOFILE = WM_CAP_START + 63
Public Const WM_CAP_SET_SEQUENCE_SETUP = WM_CAP_START + 64
Public Const WM_CAP_GET_SEQUENCE_SETUP = WM_CAP_START + 65
Public Const WM_CAP_SET_MCI_DEVICE = WM_CAP_START + 66
Public Const WM_CAP_GET_MCI_DEVICE = WM_CAP_START + 67
Public Const WM_CAP_STOP = WM_CAP_START + 68
Public Const WM_CAP_ABORT = WM_CAP_START + 69

Public Const WM_CAP_SINGLE_FRAME_OPEN = WM_CAP_START + 70
Public Const WM_CAP_SINGLE_FRAME_CLOSE = WM_CAP_START + 71
Public Const WM_CAP_SINGLE_FRAME = WM_CAP_START + 72

Public Const WM_CAP_PAL_OPEN = WM_CAP_START + 80
Public Const WM_CAP_PAL_SAVE = WM_CAP_START + 81
Public Const WM_CAP_PAL_PASTE = WM_CAP_START + 82
Public Const WM_CAP_PAL_AUTOCREATE = WM_CAP_START + 83
Public Const WM_CAP_PAL_MANUALCREATE = WM_CAP_START + 84

'// Post agregado de siguiente VFW 1.1
Public Const WM_CAP_SET_CALLBACK_CAPCONTROL = WM_CAP_START + 85

'// Definir el mensaje de captura
Public Const WM_CAP_END = WM_CAP_SET_CALLBACK_CAPCONTROL

'// ------------------------------------------------------------------
'// Estructuras
'// ------------------------------------------------------------------
Type CAPDRIVERCAPS
    wDeviceIndex As Long '               // Índice de ConvDriver en system.ini
    fHasOverlay As Long '                // ¿Puede el dispositivo sobreponer?
    fHasDlgVideoSource As Long '         // ¿Tiene fuente video dlg?
    fHasDlgVideoFormat As Long '         // ¿Tiene formato de video dlg?
    fHasDlgVideoDisplay As Long '        // ¿Tiene dlg externo?
    fCaptureInitialized As Long '        // ¿El dispositivo permite captura?
    fDriverSuppliesPalettes As Long '    // ¿El dispositivo puede crear paletas?
    hVideoIn As Long '                   // Dispositivo en el canal
    hVideoOut As Long '                  // Dispositivo fuera del canale
    hVideoExtIn As Long '                // Extensión del dispositivo en canal
    hVideoExtOut As Long '               // La extensión del dispositivo de salida
End Type

Type CAPSTATUS
    uiImageWidth As Long                    '// Ancho de la imagen
    uiImageHeight As Long                   '// Altura de la imagen
    fLiveWindow As Long                     '// ¿Se puede ver vista previa de video?
    fOverlayWindow As Long                  '// ¿Se puede Sobreponer?
    fScale As Long                          '// ¿Escala de imagen a cliente?
    ptScroll As POINTAPI                    '// Posición del Scroll
    fUsingDefaultPalette As Long            '// ¿Usar paleta defecto?
    fAudioHardware As Long                  '// ¿Dispositivo de audio presente?
    fCapFileExists As Long                  '// ¿Existe la fila de captura?
    dwCurrentVideoFrame As Long             '// # de cuadros de video cap'td
    dwCurrentVideoFramesDropped As Long     '// # Cuadros eliminados
    dwCurrentWaveSamples As Long            '// # Muestra de Wave
    dwCurrentTimeElapsedMS As Long          '// Poner tiempo de duracion
    hPalCurrent As Long                     '// Usar paleta actual
    fCapturingNow As Long                   '// ¿Captura en progreso?
    dwReturn As Long                        '// Valor de error en la operación
    wNumVideoAllocated As Long              '// Número actual de buffers de video
    wNumAudioAllocated As Long              '// Número actual de buffers de audio
End Type

Type CAPTUREPARMS
    dwRequestMicroSecPerFrame As Long       '// Recibir imagen
    fMakeUserHitOKToCapture As Long         '// ¿La demostración “golpeó MUY BIEN para capsular” el dlg?

    wPercentDropForError As Long            '// Mostrar si hay error > (10%)
    fYield As Long                          '// ¿Capturar a tarea de fondo?
    dwIndexSize As Long                     '// Maximo indice de tamaño de captura (32K)
    wChunkGranularity As Long               '// Granularity del pedazo de la chatarra (2K)
    fUsingDOSMemory As Long                 '// ¿Usar buffers DOS?
    wNumVideoRequested As Long              '// # video buffers, si 0, autocalc
    fCaptureAudio As Long                   '// ¿Capturar audio?
    wNumAudioRequested As Long              '// # audio buffers, si 0, autocalc
    vKeyAbort As Long                       '// Boton causa detenevión
    fAbortLeftMouse As Long                 '// Detener con el boton izquierdo
    fAbortRightMouse As Long                '// Detener con el boton derecho
    fLimitEnabled As Long                   '// Usar wTimeLimit?
    wTimeLimit As Long                      '// Seconds a capturar
    fMCIControl As Long                     '// ¿Usar recurso de Video MCI?
    fStepMCIDevice As Long                  '// ¿Cuadros al dispositivo MCI?
    dwMCIStartTime As Long                  '// Iniciar Tiempo en MS
    dwMCIStopTime As Long                   '// Detener Tiempo en MS
    fStepCaptureAt2x As Long                '// Hacer un promedio 2x
    wStepCaptureAverageFrames As Long       '// Clips temporales del promedio n
    dwAudioBufferSize As Long               '// Tamaño de bufs de audio (0 = defecto)
    fDisableWriteCache As Long              '// Detener escritura de caché
End Type

Type CAPINFOCHUNK
    fccInfoID As Long                       '// Ver Derechos
    lpData As Long                          '// posicionador de datos
    cbData As Long                          '// tamaño lpData
End Type

Type VIDEOHDR
    lpData As Long '// nombre del buffer
    dwBufferLength As Long '// Tamaño en Bits de los datos
    dwBytesUsed As Long '// ver abajo
    dwTimeCaptured As Long '// ver abajo
    dwUser As Long '// especificar datos
    dwFlags As Long '// ver abajo
    dwReserved(3) As Long '// reservado (no en uso)}
End Type

'// Hay 2 funciones exportadas a avicap32.dll
Declare Function capCreateCaptureWindowA Lib "avicap32.dll" ( _
    ByVal lpszWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, _
    ByVal hWndParent As Long, ByVal nID As Long) As Long
Declare Function capGetDriverDescriptionA Lib "avicap32.dll" ( _
    ByVal wDriver As Integer, _
    ByVal lpszName As String, _
    ByVal cbName As Long, _
    ByVal lpszVer As String, _
    ByVal cbVer As Long) As Boolean


Public Const IDS_CAP_BEGIN = 300              '/* "Inicio de Captura" */
Public Const IDS_CAP_END = 301                '/* "Fin de Captura" */

Public Const IDS_CAP_INFO = 401               '/* "%s" */
Public Const IDS_CAP_OUTOFMEM = 402           '/* "Fuera de memoria" */
Public Const IDS_CAP_FILEEXISTS = 403         '/* "Fila '%s' existe -- ¿Sobreescribir?" */
Public Const IDS_CAP_ERRORPALOPEN = 404       '/* "Error al abrir la paleta '%s'" */
Public Const IDS_CAP_ERRORPALSAVE = 405       '/* "Error al guardar la paleta '%s'" */
Public Const IDS_CAP_ERRORDIBSAVE = 406       '/* "Error al guardar el Clip '%s'" */
Public Const IDS_CAP_DEFAVIEXT = 407          '/* "avi" */
Public Const IDS_CAP_DEFPALEXT = 408          '/* "pal" */
Public Const IDS_CAP_CANTOPEN = 409           '/* "No se puede abrir '%s'" */
Public Const IDS_CAP_SEQ_MSGSTART = 410       '/* "Seleccione aceptar para iniciar la captura
Public Const IDS_CAP_SEQ_MSGSTOP = 411        '/* "Presione Esc para terminar la captura" */
                
Public Const IDS_CAP_VIDEDITERR = 412         '/* "Un error ha ocurrido al Abrir VIDEDIT." */
Public Const IDS_CAP_READONLYFILE = 413       '/* "La '%s' es de solo lectura." */
Public Const IDS_CAP_WRITEERROR = 414         '/* "No se puede escribir dentro de la fila '%s'.\n Disco lleno." */
Public Const IDS_CAP_NODISKSPACE = 415        '/* "No hay espacio para crear un archivo de la captura en el dispositivo especificado ." */
Public Const IDS_CAP_SETFILESIZE = 416        '/* "Poner tamalo a fila" */
Public Const IDS_CAP_SAVEASPERCENT = 417      '/* "Guardando como: %2ld%%  Presione Esc para terminar." */
                
Public Const IDS_CAP_DRIVER_ERROR = 418       '/* Dispositivo especifica error */

Public Const IDS_CAP_WAVE_OPEN_ERROR = 419    '/* "Error: No se puede abrir el dispositivo de sonido\n Verifique Bits por segundo, frecuencia, canales." */
Public Const IDS_CAP_WAVE_ALLOC_ERROR = 420   '/* "Error: Fuera de memoria para buffers de wave ." */
Public Const IDS_CAP_WAVE_PREPARE_ERROR = 421 '/* "Error: No se puede preparar buffers de wave ." */
Public Const IDS_CAP_WAVE_ADD_ERROR = 422     '/* "Error: No se puede añadir buffers de wave." */
Public Const IDS_CAP_WAVE_SIZE_ERROR = 423    '/* "Error: Mal tamaño de wave." */
                
Public Const IDS_CAP_VIDEO_OPEN_ERROR = 424   '/* "Error: No se puede entrar a la entrada del dispositivo." */
Public Const IDS_CAP_VIDEO_ALLOC_ERROR = 425  '/* "Error: Fuera de memoria para los buffers de video ." */
Public Const IDS_CAP_VIDEO_PREPARE_ERROR = 426 '/*"Error: No se puede preparar los Buffers de captura." */
Public Const IDS_CAP_VIDEO_ADD_ERROR = 427    '/* "Error: No se puede añadir buffers de video." */
Public Const IDS_CAP_VIDEO_SIZE_ERROR = 428   '/* "Error: Mal tamaño de Video." */
                
Public Const IDS_CAP_FILE_OPEN_ERROR = 429    '/* "Error: No se puede capturar la fila." */
Public Const IDS_CAP_FILE_WRITE_ERROR = 430   '/* "Error: No se puede capturar. Disco lleno." */
Public Const IDS_CAP_RECORDING_ERROR = 431    '/* "Error: No se puede escribi en disco. Tamaño de datos muy grandes o disco lleno." */
Public Const IDS_CAP_RECORDING_ERROR2 = 432   '/* "Error al grabar" */
Public Const IDS_CAP_AVI_INIT_ERROR = 433     '/* "Error: No se puede iniciar la captura." */
Public Const IDS_CAP_NO_FRAME_CAP_ERROR = 434 '/* "Advertencia: No se ha capturado
Public Const IDS_CAP_NO_PALETTE_WARN = 435    '/* "Advertencia: Usando paleta defecto." */
Public Const IDS_CAP_MCI_CONTROL_ERROR = 436  '/* "Error: No se puede acceder al dispositivo MCI." */
Public Const IDS_CAP_MCI_CANT_STEP_ERROR = 437 '/*" Error: No se puede poner cuadros al dispositivo MCI." */
Public Const IDS_CAP_NO_AUDIO_CAP_ERROR = 438 '/* "Error: No hay datos de audio capturados.\n Verifique la configuración de la tarjeta de sonido." */
Public Const IDS_CAP_AVI_DRAWDIB_ERROR = 439  '/* "Error: No se puede iniciar el formato." */
Public Const IDS_CAP_COMPRESSOR_ERROR = 440   '/* "Error: No se puede iniciar el compresor." */
Public Const IDS_CAP_AUDIO_DROP_ERROR = 441   '/* "Error: No hay datos de audio, reduzca la velocidad de captura." */
                
'/* status string IDs */
Public Const IDS_CAP_STAT_LIVE_MODE = 500      '/* "Ventana activa" */
Public Const IDS_CAP_STAT_OVERLAY_MODE = 501   '/* "Ventana Superpuesta" */
Public Const IDS_CAP_STAT_CAP_INIT = 502       '/* "Configurando la captura - espere porfavor" */
Public Const IDS_CAP_STAT_CAP_FINI = 503       '/* "Fin de captura, Cuadro escrito %ld" */
Public Const IDS_CAP_STAT_PALETTE_BUILD = 504  '/* "Creando mapa de Bits" */
Public Const IDS_CAP_STAT_OPTPAL_BUILD = 505   '/* "Ver paleta*/
Public Const IDS_CAP_STAT_I_FRAMES = 506       '/* "%d Cuadros" */
Public Const IDS_CAP_STAT_L_FRAMES = 507       '/* "%ld Cuadros" */
Public Const IDS_CAP_STAT_CAP_L_FRAMES = 508   '/* "Capturado %ld Cuadros" */
Public Const IDS_CAP_STAT_CAP_AUDIO = 509      '/* "Capturando audio" */
Public Const IDS_CAP_STAT_VIDEOCURRENT = 510   '/* "Capturado %ld Cuadros (%ld Eliminados) %d.%03d sec." */
Public Const IDS_CAP_STAT_VIDEOAUDIO = 511     '/* "Capturado %d.%03d sec.  %ld Cuadros (%ld Eliminados) (%d.%03d fps).  %ld bytes audio(%d,%03d sps)" */
Public Const IDS_CAP_STAT_VIDEOONLY = 512      '/* "Captured %d.%03d sec.  %ld Cuadros (%ld Eliminados) (%d.%03d fps)" */
Function capSetCallbackOnError(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnError = SendMessage(lwnd, WM_CAP_SET_CALLBACK_ERROR, 0, lpProc)
End Function
Function capSetCallbackOnStatus(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnStatus = SendMessage(lwnd, WM_CAP_SET_CALLBACK_STATUS, 0, lpProc)
End Function
Function capSetCallbackOnYield(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnYield = SendMessage(lwnd, WM_CAP_SET_CALLBACK_YIELD, 0, lpProc)
End Function
Function capSetCallbackOnFrame(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnFrame = SendMessage(lwnd, WM_CAP_SET_CALLBACK_FRAME, 0, lpProc)
End Function
Function capSetCallbackOnVideoStream(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnVideoStream = SendMessage(lwnd, WM_CAP_SET_CALLBACK_VIDEOSTREAM, 0, lpProc)
End Function
Function capSetCallbackOnWaveStream(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnWaveStream = SendMessage(lwnd, WM_CAP_SET_CALLBACK_WAVESTREAM, 0, lpProc)
End Function
Function capSetCallbackOnCapControl(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnCapControl = SendMessage(lwnd, WM_CAP_SET_CALLBACK_CAPCONTROL, 0, lpProc)
End Function
Function capSetUserData(ByVal lwnd As Long, ByVal lUser As Long) As Boolean
   capSetUserData = SendMessage(lwnd, WM_CAP_SET_USER_DATA, 0, lUser)
End Function
Function capGetUserData(ByVal lwnd As Long) As Long
   capGetUserData = SendMessage(lwnd, WM_CAP_GET_USER_DATA, 0, 0)
End Function
Function capDriverConnect(ByVal lwnd As Long, ByVal i As Integer) As Boolean
   capDriverConnect = SendMessage(lwnd, WM_CAP_DRIVER_CONNECT, i, 0)
End Function
Function capDriverDisconnect(ByVal lwnd As Long) As Boolean
   capDriverDisconnect = SendMessage(lwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0)
End Function
Function capDriverGetName(ByVal lwnd As Long, ByVal szName As Long, ByVal wSize As Integer) As Boolean
   capDriverGetName = SendMessage(lwnd, YOURCONSTANTMESSAGE, wSize, szName)
End Function
Function capDriverGetVersion(ByVal lwnd As Long, ByVal szVer As Long, ByVal wSize As Integer) As Boolean
   capDriverGetVersion = SendMessage(lwnd, WM_CAP_DRIVER_GET_VERSION, wSize, szVer)
End Function
Function capDriverGetCaps(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Boolean
   capDriverGetCaps = SendMessage(lwnd, WM_CAP_DRIVER_GET_CAPS, wSize, s)
End Function
Function capFileSetCaptureFile(ByVal lwnd As Long, szName As String) As Boolean
   capFileSetCaptureFile = SendMessageS(lwnd, WM_CAP_FILE_SET_CAPTURE_FILE, 0, szName)
End Function
Function capFileGetCaptureFile(ByVal lwnd As Long, ByVal szName As Long, wSize As String) As Boolean
   capFileGetCaptureFile = SendMessageS(lwnd, WM_CAP_FILE_SET_CAPTURE_FILE, wSize, szName)
End Function
Function capFileAlloc(ByVal lwnd As Long, ByVal dwSize As Long) As Boolean
   capFileAlloc = SendMessage(lwnd, WM_CAP_FILE_ALLOCATE, 0, dwSize)
End Function
Function capFileSaveAs(ByVal lwnd As Long, szName As String) As Boolean
   capFileSaveAs = SendMessageS(lwnd, WM_CAP_FILE_SAVEAS, 0, szName)
End Function
Function capFileSetInfoChunk(ByVal lwnd As Long, ByVal lpInfoChunk As Long) As Boolean
   capFileSetInfoChunk = SendMessage(lwnd, WM_CAP_FILE_SET_INFOCHUNK, 0, lpInfoChunk)
End Function
Function capFileSaveDIB(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capFileSaveDIB = SendMessage(lwnd, WM_CAP_FILE_SAVEDIB, 0, szName)
End Function
Function capEditCopy(ByVal lwnd As Long) As Boolean
   capEditCopy = SendMessage(lwnd, WM_CAP_EDIT_COPY, 0, 0)
End Function
Function capSetAudioFormat(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Boolean
   capSetAudioFormat = SendMessage(lwnd, WM_CAP_SET_AUDIOFORMAT, wSize, s)
End Function
Function capGetAudioFormat(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Long
   capGetAudioFormat = SendMessage(lwnd, WM_CAP_GET_AUDIOFORMAT, wSize, s)
End Function
Function capGetAudioFormatSize(ByVal lwnd As Long) As Long
   capGetAudioFormatSize = SendMessage(lwnd, WM_CAP_GET_AUDIOFORMAT, 0, 0)
End Function
Function capDlgVideoFormat(ByVal lwnd As Long) As Boolean
   capDlgVideoFormat = SendMessage(lwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0)
End Function
Function capDlgVideoSource(ByVal lwnd As Long) As Boolean
   capDlgVideoSource = SendMessage(lwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0)
End Function
Function capDlgVideoDisplay(ByVal lwnd As Long) As Boolean
   capDlgVideoDisplay = SendMessage(lwnd, WM_CAP_DLG_VIDEODISPLAY, 0, 0)
End Function
Function capDlgVideoCompression(ByVal lwnd As Long) As Boolean
   capDlgVideoCompression = SendMessage(lwnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0, 0)
End Function
Function capGetVideoFormat(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Long
   capGetVideoFormat = SendMessage(lwnd, WM_CAP_GET_VIDEOFORMAT, wSize, s)
End Function
Function capGetVideoFormatSize(ByVal lwnd As Long) As Long
   capGetVideoFormatSize = SendMessage(lwnd, WM_CAP_GET_VIDEOFORMAT, 0, 0)
End Function
Function capSetVideoFormat(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Boolean
   capSetVideoFormat = SendMessage(lwnd, WM_CAP_SET_VIDEOFORMAT, wSize, s)
End Function
Function capPreview(ByVal lwnd As Long, ByVal f As Boolean) As Boolean
   capPreview = SendMessage(lwnd, WM_CAP_SET_PREVIEW, f, 0)
End Function
Function capPreviewRate(ByVal lwnd As Long, ByVal wMS As Integer) As Boolean
   capPreviewRate = SendMessage(lwnd, WM_CAP_SET_PREVIEWRATE, wMS, 0)
End Function
Function capOverlay(ByVal lwnd As Long, ByVal f As Boolean) As Boolean
   capOverlay = SendMessage(lwnd, WM_CAP_SET_OVERLAY, f, 0)
End Function
Function capPreviewScale(ByVal lwnd As Long, ByVal f As Boolean) As Boolean
   capPreviewScale = SendMessage(lwnd, WM_CAP_SET_SCALE, f, 0)
End Function
'Function capGetStatus(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Boolean
'   capGetStatus = SendMessage(lwnd, WM_CAP_GET_STATUS, wSize, s)
'End Function
Function capGetStatus(ByVal hCapWnd As Long, ByRef capStat As CAPSTATUS) As Boolean
   capGetStatus = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)
End Function
Function capSetScrollPos(ByVal lwnd As Long, ByVal lpP As Long) As Boolean
   capSetScrollPos = SendMessage(lwnd, WM_CAP_SET_SCROLL, 0, lpP)
End Function
Function capGrabFrame(ByVal lwnd As Long) As Boolean
   capGrabFrame = SendMessage(lwnd, WM_CAP_GRAB_FRAME, 0, 0)
End Function
Function capGrabFrameNoStop(ByVal lwnd As Long) As Boolean
   capGrabFrameNoStop = SendMessage(lwnd, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0)
End Function
Function capCaptureSequence(ByVal lwnd As Long) As Boolean
   capCaptureSequence = SendMessage(lwnd, WM_CAP_SEQUENCE, 0, 0)
End Function
Function capCaptureSequenceNoFile(ByVal lwnd As Long) As Boolean
   capCaptureSequenceNoFile = SendMessage(lwnd, WM_CAP_SEQUENCE_NOFILE, 0, 0)
End Function
Function capCaptureStop(ByVal lwnd As Long) As Boolean
   capCaptureStop = SendMessage(lwnd, WM_CAP_STOP, 0, 0)
End Function
Function capCaptureAbort(ByVal lwnd As Long) As Boolean
   capCaptureAbort = SendMessage(lwnd, WM_CAP_ABORT, 0, 0)
End Function
Function capCaptureSingleFrameOpen(ByVal lwnd As Long) As Boolean
   capCaptureSingleFrameOpen = SendMessage(lwnd, WM_CAP_SINGLE_FRAME_OPEN, 0, 0)
End Function
Function capCaptureSingleFrameClose(ByVal lwnd As Long) As Boolean
   capCaptureSingleFrameClose = SendMessage(lwnd, WM_CAP_SINGLE_FRAME_CLOSE, 0, 0)
End Function
Function capCaptureSingleFrame(ByVal lwnd As Long) As Boolean
   capCaptureSingleFrame = SendMessage(lwnd, WM_CAP_SINGLE_FRAME, 0, 0)
End Function
Function capCaptureGetSetup(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Boolean
   capCaptureGetSetup = SendMessage(lwnd, WM_CAP_GET_SEQUENCE_SETUP, wSize, s)
End Function
Function capCaptureSetSetup(ByVal lwnd As Long, ByVal s As Long, ByVal wSize As Integer) As Boolean
   capCaptureSetSetup = SendMessage(lwnd, WM_CAP_SET_SEQUENCE_SETUP, wSize, s)
End Function
Function capSetMCIDeviceName(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capSetMCIDeviceName = SendMessage(lwnd, WM_CAP_SET_MCI_DEVICE, 0, szName)
End Function
Function capGetMCIDeviceName(ByVal lwnd As Long, ByVal szName As Long, ByVal wSize As Integer) As Boolean
   capGetMCIDeviceName = SendMessage(lwnd, WM_CAP_GET_MCI_DEVICE, wSize, szName)
End Function
Function capPaletteOpen(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capPaletteOpen = SendMessage(lwnd, WM_CAP_PAL_OPEN, 0, szName)
End Function
Function capPaletteSave(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capPaletteSave = SendMessage(lwnd, WM_CAP_PAL_SAVE, 0, szName)
End Function
Function capPalettePaste(ByVal lwnd As Long) As Boolean
   capPalettePaste = SendMessage(lwnd, WM_CAP_PAL_PASTE, 0, 0)
End Function
Function capPaletteAuto(ByVal lwnd As Long, ByVal iFrames As Integer, ByVal iColor As Long) As Boolean
   capPaletteAuto = SendMessage(lwnd, WM_CAP_PAL_AUTOCREATE, iFrames, iColors)
End Function
Function capPaletteManual(ByVal lwnd As Long, ByVal fGrab As Boolean, ByVal iColors As Long) As Boolean
   capPaletteManual = SendMessage(lwnd, WM_CAP_PAL_MANUALCREATE, fGrab, iColors)
End Function

