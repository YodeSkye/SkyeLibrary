
Namespace Contracts

    ''' <summary>
    ''' Defines a common contract for media playback engines.
    ''' Provides methods for controlling playback, accessing media properties,
    ''' and receiving events about playback state.
    ''' </summary>
    Public Interface IMediaPlayer

        ''' <summary>
        ''' Begins playback of the specified media file from a file system path.
        ''' </summary>
        ''' <param name="path">The full file path of the media to play.</param>
        Sub Play(path As String)
        ''' <summary>
        ''' Begins playback of the specified media from a URI.
        ''' </summary>
        ''' <param name="uri">The URI of the media to play (local or remote).</param>
        Sub Play(uri As Uri)
        ''' <summary>
        ''' Resumes playback of the current media file.
        ''' </summary>
        Sub Play()
        ''' <summary>
        ''' Pauses playback if media is currently playing.
        ''' </summary>
        Sub Pause()
        ''' <summary>
        ''' Stops playback. Implementations may retain the loaded media
        ''' for later replay, or clear it depending on design.
        ''' </summary>
        Sub [Stop]()

        ''' <summary>
        ''' Indicates whether a media item is currently loaded in the player.
        ''' </summary>
        ReadOnly Property HasMedia As Boolean
        ''' <summary>
        ''' Gets the file path or URI of the currently loaded media, if available.
        ''' </summary>
        ReadOnly Property Path As String
        ''' <summary>
        ''' Gets or sets the playback volume as an integer percentage (0–100).
        ''' </summary>
        Property Volume As Integer
        ''' <summary>
        ''' Gets or sets the current playback position in seconds.
        ''' Values typically range from 0.0 to <see cref="Duration"/>.
        ''' </summary>
        Property Position As Double
        ''' <summary>
        ''' Gets the total duration of the current media in seconds.
        ''' Returns 0 if no media is loaded or duration is unknown.
        ''' </summary>
        ReadOnly Property Duration As Double
        ''' <summary>
        ''' Gets the width of the video in pixels, if applicable.
        ''' Returns 0 for audio-only media.
        ''' </summary>
        ReadOnly Property VideoWidth As Integer
        ''' <summary>
        ''' Gets the height of the video in pixels, if applicable.
        ''' Returns 0 for audio-only media.
        ''' </summary>
        ReadOnly Property VideoHeight As Integer
        ''' <summary>
        ''' Gets the aspect ratio of the video as width ÷ height.
        ''' Returns 0 if not applicable.
        ''' </summary>
        ReadOnly Property AspectRatio As Double

        ''' <summary>
        ''' Raised when playback starts at the beginning of the current media.
        ''' </summary>
        Event PlaybackStarted()
        ''' <summary>
        ''' Raised when playback reaches the end of the current media.
        ''' </summary>
        Event PlaybackEnded()

    End Interface

End Namespace
