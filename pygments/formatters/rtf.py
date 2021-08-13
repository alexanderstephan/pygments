"""
    pygments.formatters.rtf
    ~~~~~~~~~~~~~~~~~~~~~~~

    A formatter that generates RTF files.

    :copyright: Copyright 2006-2021 by the Pygments team, see AUTHORS.
    :license: BSD, see LICENSE for details.
"""

from pygments.formatter import Formatter
from pygments.util import get_int_opt, surrogatepair
from pygments.token import Token


__all__ = ['RtfFormatter']


class RtfFormatter(Formatter):
    """
    Format tokens as RTF markup. This formatter automatically outputs full RTF
    documents with color information and other useful stuff. Perfect for Copy and
    Paste into Microsoft(R) Word(R) documents.

    Please note that ``encoding`` and ``outencoding`` options are ignored.
    The RTF format is ASCII natively, but handles unicode characters correctly
    thanks to escape sequences.

    .. versionadded:: 0.6

    Additional options accepted:

    `style`
        The style to use, can be a string or a Style subclass (default:
        ``'default'``).

    `fontface`
        The used font family, for example ``Bitstream Vera Sans``. Defaults to
        some generic font which is supposed to have fixed width.

    `fontsize`
        Size of the font used. Size is specified in half points. The
        default is 24 half-points, giving a size 12 font.

    `linenos`
        Turn on line numbering, if set to anything but ``'0'`` 
        (default: ``'0'``).

    `lineno_fontsize`
        Font size for line numbers. Size is specified in half points. The
        default is 18 half-points, giving a 9pt font.

    `linenostart`
        The line number for the first line (default: ``1``).

    `linenostep`
        If set to a number n > 1, only every nth line number is printed.

        .. versionadded:: 2.0
    """
    name = 'RTF'
    aliases = ['rtf']
    filenames = ['*.rtf']

    def __init__(self, **options):
        r"""
        Additional options accepted:

        ``fontface``
            Name of the font used. Could for example be ``'Courier New'``
            to further specify the default which is ``'\fmodern'``. The RTF
            specification claims that ``\fmodern`` are "Fixed-pitch serif
            and sans serif fonts". Hope every RTF implementation thinks
            the same about modern...

        """
        Formatter.__init__(self, **options)
        self.fontface = options.get('fontface') or ''
        self.fontsize = get_int_opt(options, 'fontsize', 0)
        self.linenos = options.get('linenos', False)
        self.lineno_fontsize = get_int_opt(options, 'lineno_fontsize', 18)
        self.linenostart = abs(get_int_opt(options, 'linenostart', 1))
        self.linenostep = abs(get_int_opt(options, 'linenostep', 1))

    def _escape(self, text):
        return text.replace('\\', '\\\\') \
                   .replace('{', '\\{') \
                   .replace('}', '\\}')

    def _escape_text(self, text):
        # empty strings, should give a small performance improvement
        if not text:
            return ''

        # escape text
        text = self._escape(text)

        buf = []
        for c in text:
            cn = ord(c)
            if cn < (2**7):
                # ASCII character
                buf.append(str(c))
            elif (2**7) <= cn < (2**16):
                # single unicode escape sequence
                buf.append('{\\u%d}' % cn)
            elif (2**16) <= cn:
                # RTF limits unicode to 16 bits.
                # Force surrogate pairs
                buf.append('{\\u%d}{\\u%d}' % surrogatepair(cn))

        return ''.join(buf).replace('\n', '\\par\n')

    def format_unencoded(self, tokensource, outfile):
        # rtf 1.8 header
        outfile.write('{\\rtf1\\ansi\\uc0\\deff0'
                      '{\\fonttbl{\\f0\\fmodern\\fprq1\\fcharset0%s;}}'
                      '{\\colortbl;' % (self.fontface and
                                        ' ' + self._escape(self.fontface) or
                                        ''))

        # convert colors and save them in a mapping to access them later.
        color_mapping = {}
        offset = 1
        for _, style in self.style:
            for color in style['color'], style['bgcolor'], style['border']:
                if color and color not in color_mapping:
                    color_mapping[color] = offset
                    outfile.write('\\red%d\\green%d\\blue%d;' % (
                        int(color[0:2], 16),
                        int(color[2:4], 16),
                        int(color[4:6], 16)
                    ))
                    offset += 1
        outfile.write('}\\f0 ')
        if self.fontsize:
            outfile.write('\\fs%d' % self.fontsize)

        ## line numbering setup
        if self.linenos:
            
            # first pass of tokens (tokensource) to obtain total number of lines 
            # and deal with tokens of type Literal.String.Doc
            num_of_lines = 0
            tokens = []
            for ttype, value in tokensource:

                # tokensource is a generator, so make a copy
                if ttype == Token.Literal.String.Doc:
                    lines = value.split('\n')
                    for line in lines:
                        num_of_lines += 1
                        tokens.append((ttype, line+'\n'))
                    
                elif value == u'\n':
                    num_of_lines += 1
                    tokens.append((ttype, value))
                else:
                    tokens.append((ttype, value))
                    

            # width of (space) padded line number string
            linenos_width = len(str(num_of_lines))            

            # line number output book keeping
            lineno = self.linenostart
            output_lineno = True

        else:
            tokens = tokensource
        
        
        # highlight stream
        for ttype, value in tokens:
            while not self.style.styles_token(ttype) and ttype.parent:
                ttype = ttype.parent
            style = self.style.style_for_token(ttype)
            buf = []

            # line numbers
            if self.linenos:                                 

                if output_lineno:                                      

                    if (lineno-self.linenostep)%self.linenostep == 0:
                        lineno_str = str(lineno).rjust(linenos_width)      
                    else:
                        lineno_str = "".rjust(linenos_width)      

                    outfile.write(u'{\\fs%d \\cf1 %s  }'\
                        % (self.lineno_fontsize, lineno_str))          
                    output_lineno = False                              
                                                                       
                if value.endswith('\n'):
                    output_lineno = True                               
                    lineno += 1                                        

            if style['bgcolor']:
                buf.append('\\cb%d' % color_mapping[style['bgcolor']])
            if style['color']:
                buf.append('\\cf%d' % color_mapping[style['color']])
            if style['bold']:
                buf.append('\\b')
            if style['italic']:
                buf.append('\\i')
            if style['underline']:
                buf.append('\\ul')
            if style['border']:
                buf.append('\\chbrdr\\chcfpat%d' %
                           color_mapping[style['border']])
            start = ''.join(buf)

            if start:
                outfile.write('{%s ' % start)

            outfile.write(self._escape_text(value))
            if start:
                outfile.write('}')

        outfile.write('}')
