use strict;
use warnings;
use utf8;
use Text::ParseWords;
binmode(STDOUT, ':encoding(UTF-8)');

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
my $RELATORIO_CSV  = "relatorio.csv";
my $XLSX_DIR       = "xlsx_extract";   # unzipped workbook.xlsx
my $TEMPLATE_FILE  = "dashboard_template.html";
my $OUTPUT_FILE    = "dashboard_otd.generated.html";

my @MESES_PT = ('', 'Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro');
my @MESES_ABR = ('', 'jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez');

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
sub br_num {
    my ($s) = @_;
    return undef unless defined $s;
    $s =~ s/^\s+|\s+$//g;
    return undef if $s eq '' || $s eq '#REF!';
    $s =~ s/\.//g;
    $s =~ s/,/./;
    return $s + 0;
}
sub br_pct {
    my ($s) = @_;
    return undef unless defined $s;
    $s =~ s/^\s+|\s+$//g;
    return undef if $s eq '' || $s eq '#REF!';
    $s =~ s/%//;
    $s =~ s/,/./;
    return $s + 0;
}
sub fmt_pct1 {
    my ($v) = @_;
    my $s = sprintf("%.1f", $v);
    $s =~ s/\./,/;
    return $s;
}
sub read_lines_utf8 {
    my ($file) = @_;
    open(my $fh, '<:encoding(UTF-8)', $file) or die "$file: $!";
    my @lines = <$fh>;
    close $fh;
    chomp @lines;
    return @lines;
}
sub read_all_utf8 {
    my ($file) = @_;
    open(my $fh, '<:encoding(UTF-8)', $file) or die "$file: $!";
    local $/;
    my $c = <$fh>;
    close $fh;
    return $c;
}
sub xml_unescape {
    my ($s) = @_;
    return '' unless defined $s;
    $s =~ s/&lt;/</g; $s =~ s/&gt;/>/g; $s =~ s/&quot;/"/g; $s =~ s/&apos;/'/g; $s =~ s/&amp;/&/g;
    return $s;
}
# Convert a spreadsheet column letter (A, B, ..., Z, AA, AB...) to a 1-based index
sub col_to_num {
    my ($col) = @_;
    my $n = 0;
    for my $ch (split //, $col) { $n = $n * 26 + (ord(uc($ch)) - ord('A') + 1); }
    return $n;
}
sub num_to_col {
    my ($n) = @_;
    my $s = '';
    while ($n > 0) {
        my $rem = ($n - 1) % 26;
        $s = chr(ord('A') + $rem) . $s;
        $n = int(($n - 1) / 26);
    }
    return $s;
}
# nome de mes em pt-BR -> numero (aceita maiuscula/minuscula)
sub mes_num {
    my ($nome) = @_;
    return 0 unless defined $nome;
    $nome =~ s/^\s+|\s+$//g;
    for my $i (1..12) { return $i if lc($nome) eq lc($MESES_PT[$i]); }
    return 0;
}
# inteiro com separador de milhar pt-BR
sub fmt_int {
    my ($n) = @_;
    my $s = sprintf("%d", $n);
    1 while $s =~ s/(\d)(\d{3})(?!\d)/$1.$2/;
    return $s;
}
# valor em reais, 2 casas, virgula decimal
sub fmt_brl {
    my ($v) = @_;
    my ($int, $dec) = split /\./, sprintf("%.2f", $v);
    1 while $int =~ s/(\d)(\d{3})(?!\d)/$1.$2/;
    return "$int,$dec";
}
# valor pequeno (custo unitario), 3 casas
sub fmt_brl3 {
    my ($v) = @_;
    my $s = sprintf("%.3f", $v);
    $s =~ s/\./,/;
    return $s;
}

# ---------------------------------------------------------------------------
# PARSE RELATORIO (status diario) from clean CSV export
# ---------------------------------------------------------------------------
die "$RELATORIO_CSV nao encontrado\n" unless -s $RELATORIO_CSV;
my @rlines = read_lines_utf8($RELATORIO_CSV);
my ($today_y, $today_m, $today_d, $today_h, $today_min) = (gmtime(time - 3*3600))[5,4,3,2,1];
$today_y += 1900; $today_m += 1;

my %by_month;
my %pecas_mes;              # numero do mes -> total de pecas faturadas
my $total_dias = 0; my $total_verde = 0;

for my $ln (@rlines) {
    my @f = parse_line(",", 0, $ln);
    my $data = $f[1] // '';
    next unless $data =~ m{^(\d{2})/(\d{2})/(\d{4})$};
    my ($d, $m, $y) = ($1, $2, $3);
    next if ($f[2] // '') eq 'Feriado';
    my $status = $f[12] // '';
    next unless $status =~ /^(verde|vermelho)$/i;
    next if ($y > $today_y) || ($y == $today_y && $m > $today_m) || ($y == $today_y && $m == $today_m && $d > $today_d);

    # pecas faturadas do dia (coluna "Qtd. Faturada") -> acumula por mes
    my $fat = br_num($f[9]);
    $pecas_mes{$m+0} += $fat if defined $fat;

    my $key = sprintf("%04d-%02d", $y, $m);
    $by_month{$key}{name} = $MESES_PT[$m+0];
    $by_month{$key}{mnum} = $m + 0;
    push @{ $by_month{$key}{rows} }, { d => $d, g => (lc($status) eq 'verde' ? 1 : 0) };
    $total_dias++;
    $total_verde++ if lc($status) eq 'verde';
}
die "Nenhum dia valido encontrado na aba Relatorio (CSV). Abortando.\n" unless $total_dias > 0;

my @month_keys = sort keys %by_month;
# Sempre considerar somente o mes corrente, nao o historico.
my $current_month_key = sprintf("%04d-%02d", $today_y, $today_m);
@month_keys = grep { $_ eq $current_month_key } @month_keys;
$total_dias = 0; $total_verde = 0;
for my $key (@month_keys) {
    for my $r (@{ $by_month{$key}{rows} }) {
        $total_dias++;
        $total_verde++ if $r->{g};
    }
}
die "Nenhum dia do mes corrente ($current_month_key) encontrado na aba Relatorio. Abortando.\n" unless $total_dias > 0;
my $pct_vermelho_total = 100 * ($total_dias - $total_verde) / $total_dias;

# ---------------------------------------------------------------------------
# LOCATE "Metas" SHEET INSIDE THE UNZIPPED XLSX
# ---------------------------------------------------------------------------
my $wb_xml = read_all_utf8("$XLSX_DIR/xl/workbook.xml");
my $rels_xml = read_all_utf8("$XLSX_DIR/xl/_rels/workbook.xml.rels");

my ($metas_rid) = $wb_xml =~ /<sheet[^>]*name="Metas"[^>]*r:id="(rId\d+)"/;
die "Nao encontrei a aba 'Metas' em workbook.xml.\n" unless $metas_rid;

my ($metas_target) = $rels_xml =~ /<Relationship Id="\Q$metas_rid\E"[^>]*Target="([^"]+)"/;
die "Nao encontrei o alvo do relationship $metas_rid.\n" unless $metas_target;
$metas_target =~ s{^/}{};
my $metas_sheet_file = "$XLSX_DIR/xl/$metas_target";
die "$metas_sheet_file nao existe.\n" unless -s $metas_sheet_file;

# shared strings table (index => text)
my @shared;
if (-s "$XLSX_DIR/xl/sharedStrings.xml") {
    my $ss_xml = read_all_utf8("$XLSX_DIR/xl/sharedStrings.xml");
    while ($ss_xml =~ /<si>(.*?)<\/si>/gs) {
        my $si = $1;
        my $text = join('', ($si =~ /<t[^>]*>(.*?)<\/t>/gs));
        push @shared, xml_unescape($text);
    }
}

my $sheet_xml = read_all_utf8($metas_sheet_file);

# ---------------------------------------------------------------------------
# ABA "Transportes": gasto mensal com frete + volumes expedidos
# ---------------------------------------------------------------------------
# Localiza o arquivo de uma aba pelo nome (mesma logica usada para "Metas").
sub sheet_file_for {
    my ($name) = @_;
    my ($rid) = $wb_xml =~ /<sheet[^>]*name="\Q$name\E"[^>]*r:id="(rId\d+)"/;
    return undef unless $rid;
    my ($target) = $rels_xml =~ /<Relationship Id="\Q$rid\E"[^>]*Target="([^"]+)"/;
    return undef unless $target;
    $target =~ s{^/}{};
    my $path = "$XLSX_DIR/xl/$target";
    return (-s $path) ? $path : undef;
}
# Localiza uma aba cujo nome COMECA com $prefix (a aba de transporte ja foi
# renomeada de "Transportes" para "Transportes CIF"; casar so o nome exato
# fazia a secao de logistica sumir sem aviso).
sub sheet_file_starting_with {
    my ($prefix) = @_;
    while ($wb_xml =~ /<sheet[^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"/g) {
        my ($nome, $rid) = (xml_unescape($1), $2);
        next unless $nome =~ /^\Q$prefix\E/i;
        my ($target) = $rels_xml =~ /<Relationship Id="\Q$rid\E"[^>]*Target="([^"]+)"/;
        next unless $target;
        $target =~ s{^/}{};
        my $path = "$XLSX_DIR/xl/$target";
        return ($path, $nome) if -s $path;
    }
    return (undef, undef);
}
# Le uma planilha inteira para um hash {linha}{coluna} = texto ja resolvido.
sub parse_grid {
    my ($file) = @_;
    my $xml = read_all_utf8($file);
    my %g;
    while ($xml =~ /<row r="(\d+)"[^>]*>(.*?)<\/row>/gs) {
        my ($rn, $inner) = ($1, $2);
        while ($inner =~ /<c r="([A-Z]+)$rn"([^>]*)(?:\/>|>(.*?)<\/c>)/gs) {
            my ($col, $attrs, $body) = ($1, $2, $3);
            next unless defined $body;
            my ($v) = $body =~ /<v>(.*?)<\/v>/s;
            next unless defined $v;
            $g{$rn}{$col} = ($attrs =~ /t="s"/) ? ($shared[$v] // '') : xml_unescape($v);
        }
    }
    return \%g;
}

my %gasto_mes;      # mes -> gasto com transporte (R$)
my %vol_mes;        # mes -> volumes expedidos
my %txt_mes;        # mes -> qtd de valores gravados como texto (nao entram no total)
my %emp_mes;        # mes -> {transportadora => qtd de lancamentos}
my ($transp_file, $transp_nome) = sheet_file_starting_with("Transportes");
$transp_nome = 'Transportes' unless defined $transp_nome;
if ($transp_file) {
    my $tg = parse_grid($transp_file);
    # Bloco superior: coluna L = nome do mes, coluna M = valor gasto no mes.
    for my $r (1 .. 20) {
        next unless $tg->{$r};
        my $mes = $tg->{$r}{L};
        my $val = $tg->{$r}{M};
        next unless defined $mes && defined $val;
        my $mi = mes_num($mes);
        next unless $mi;
        next unless $val =~ /^-?[\d.]+$/;
        $gasto_mes{$mi} = $val + 0;
    }
    # Log detalhado (linha 15 em diante): J = mes, F = volumes, I = valor.
    for my $r (sort { $a <=> $b } keys %$tg) {
        next if $r < 15;
        my $mi = mes_num($tg->{$r}{J});
        next unless $mi;
        my $v = $tg->{$r}{F};
        $vol_mes{$mi} += $v if defined $v && $v =~ /^\d+(?:\.\d+)?$/;
        # valor digitado como texto (ex: "218;89") nao entra no total da planilha
        my $val = $tg->{$r}{I};
        $txt_mes{$mi}++ if defined $val && $val ne '' && $val !~ /^-?[\d.]+$/
                        && $val !~ /^(coleta|full|Full|Lan.ar Manual)$/ && !mes_num($val);
        my $e = $tg->{$r}{G};
        if (defined $e) {
            $e =~ s/^\s+|\s+$//g;
            $emp_mes{$mi}{ucfirst(lc $e)}++
                if $e ne '' && $e ne '-' && $e ne ',' && $e ne 'Empresa' && $e !~ /^\*+$/;
        }
    }
}

# Janela de logistica: os ultimos 3 meses que tem gasto de transporte lancado.
my @log_months = sort { $a <=> $b } grep { ($gasto_mes{$_} // 0) > 0 } keys %gasto_mes;
@log_months = @log_months[-3 .. -1] if scalar(@log_months) > 3;

sub get_cell_raw {
    my ($xml, $ref) = @_;
    if ($xml =~ /<c r="\Q$ref\E"([^>]*)>(.*?)<\/c>/s) {
        my ($attrs, $inner) = ($1, $2);
        my ($v) = $inner =~ /<v>(.*?)<\/v>/s;
        return (undef, undef) unless defined $v;
        my $is_shared = ($attrs =~ /t="s"/) ? 1 : 0;
        return ($v, $is_shared);
    }
    if ($xml =~ /<c r="\Q$ref\E"[^>]*\/>/) { return (undef, undef); }
    return (undef, undef);
}
sub get_cell_text {
    my ($xml, $ref) = @_;
    my ($v, $is_shared) = get_cell_raw($xml, $ref);
    return undef unless defined $v;
    return $is_shared ? $shared[$v] : xml_unescape($v);
}
sub get_row_xml {
    my ($xml, $rownum) = @_;
    if ($xml =~ /<row r="\Q$rownum\E"[^>]*>(.*?)<\/row>/s) {
        return $1;
    }
    return '';
}

# Row 1 of the Metas sheet: merged month-name header cells across many columns.
# Scan columns C..BZ looking for cells whose resolved text matches a known month name;
# the Qtd.Pedidos column for that month is that cell's column, Realizado is the next one.
my $row1 = get_row_xml($sheet_xml, 1);
my @months_all;
my %seen_month;
for (my $colnum = 3; $colnum <= 78; $colnum++) {
    my $col = num_to_col($colnum);
    my $txt = get_cell_text($row1, "$col" . "1");
    next unless defined $txt;
    for my $mi (1..12) {
        if (lc($txt) eq lc($MESES_PT[$mi]) && !$seen_month{$mi}) {
            $seen_month{$mi} = 1;
            push @months_all, { name => $MESES_PT[$mi], mnum => $mi, qtd_col => $col, real_col => num_to_col($colnum+1) };
            last;
        }
    }
}
die "Nao encontrei nomes de meses na linha 1 da aba Metas.\n" unless @months_all;
@months_all = sort { $a->{mnum} <=> $b->{mnum} } @months_all;

# Rows 3..6 = Ate 2 / Ate 5 / Ate 7 / Acima 7 (labels in column A), meta value in column B
my %faixa_row;
for my $r (3..6) {
    my $rxml = get_row_xml($sheet_xml, $r);
    my $label = get_cell_text($rxml, "A$r");
    next unless defined $label;
    $label =~ s/^\s+|\s+$//g;
    $faixa_row{$label} = $rxml if $label =~ /^(Até 2|Até 5|Até 7|Acima 7)$/;
}
die "Nao encontrei as 4 linhas de faixa de prazo (Ate 2/5/7, Acima 7) na aba Metas (xlsx). Layout pode ter mudado.\n"
    unless scalar(keys %faixa_row) == 4;

my ($meta_ate2_raw) = get_cell_text($faixa_row{'Até 2'}, "B3");
# Metas oficiais por faixa de prazo, definidas pelo usuario (B3 da aba Metas traz 700%, que e um erro de digitacao).
my $meta_ate2   = 35;
my $meta_ate5   = 50;
my $meta_ate7   = 15;
my $meta_acima7 = 0;

my @active_months;
for my $m (@months_all) {
    my $total = 0;
    my %qtd;
    for my $lbl ('Até 2','Até 5','Até 7','Acima 7') {
        my $rownum = ($lbl eq 'Até 2') ? 3 : ($lbl eq 'Até 5') ? 4 : ($lbl eq 'Até 7') ? 5 : 6;
        my $qtd_ref = $m->{qtd_col} . $rownum;
        my $v = get_cell_text($faixa_row{$lbl}, $qtd_ref);
        $qtd{$lbl} = (defined $v && $v ne '') ? ($v + 0) : 0;
        $total += $qtd{$lbl};
    }
    my %row = (name => $m->{name}, mnum => $m->{mnum}, total => $total);
    for my $lbl ('Até 2','Até 5','Até 7','Acima 7') {
        my $rownum = ($lbl eq 'Até 2') ? 3 : ($lbl eq 'Até 5') ? 4 : ($lbl eq 'Até 7') ? 5 : 6;
        my $real_ref = $m->{real_col} . $rownum;
        my $v = get_cell_text($faixa_row{$lbl}, $real_ref);
        $row{$lbl} = (defined $v && $v ne '') ? ($v + 0) * 100 : 0;
    }
    push @active_months, \%row;
}
# Sempre considerar somente o mes corrente, nao o historico.
@active_months = grep { $_->{mnum} == $today_m } @active_months;
unless (@active_months) {
    # Aba Metas ainda nao tem coluna para o mes corrente (ex: inicio do mes, coluna nao criada ainda).
    # Mostrar zerado em vez de falhar todo dia ate alguem criar a coluna manualmente.
    @active_months = ({ name => $MESES_PT[$today_m], mnum => $today_m, total => 0,
        'Até 2' => 0, 'Até 5' => 0, 'Até 7' => 0, 'Acima 7' => 0 });
}

# ---------------------------------------------------------------------------
# KPI COMPUTATIONS
# ---------------------------------------------------------------------------
my $last = $active_months[-1];

my $kpi_ate2_last   = $last->{'Até 2'};
my $kpi_ate5_last   = $last->{'Até 5'};
my $kpi_ate7_last   = $last->{'Até 7'};
my $kpi_acima7_last = $last->{'Acima 7'};

# "Ate 2 dias": valor maior e melhor, meta e um piso (>=35%). gap = realizado - meta.
sub gap_html_higher_is_better {
    my ($value, $meta) = @_;
    my $gap = $value - $meta;
    return $gap >= 0
        ? qq{<div class="delta good">&#9650; } . fmt_pct1($gap) . qq{ p.p. acima da meta</div>}
        : qq{<div class="delta bad">&#9660; } . fmt_pct1(-$gap) . qq{ p.p. abaixo da meta</div>};
}
# "Ate 5", "Ate 7" e "Acima de 7 dias": valor menor e melhor, meta e um teto. gap = meta - realizado.
sub gap_html_lower_is_better {
    my ($value, $meta) = @_;
    my $gap = $meta - $value;
    return $gap >= 0
        ? qq{<div class="delta good">&#9650; dentro da meta</div>}
        : qq{<div class="delta bad">&#9660; } . fmt_pct1(-$gap) . qq{ p.p. acima da meta</div>};
}

my $kpi_ate2_gap_html   = gap_html_higher_is_better($kpi_ate2_last, $meta_ate2);
my $kpi_ate5_gap_html   = gap_html_lower_is_better($kpi_ate5_last, $meta_ate5);
my $kpi_ate7_gap_html   = gap_html_lower_is_better($kpi_ate7_last, $meta_ate7);
my $kpi_acima7_gap_html = gap_html_lower_is_better($kpi_acima7_last, $meta_acima7);

my $periodo_label = @active_months == 1
    ? $active_months[0]{name}
    : $active_months[0]{name} . "\x{2013}" . $active_months[-1]{name};

# ---------------------------------------------------------------------------
# SVG: chart 1 - distribuicao de pedidos por faixa de prazo
# ---------------------------------------------------------------------------
sub bucket_chart_svg {
    my (@months) = @_;
    my $n = scalar(@months);
    my $row_h = 42;
    my $h = 10 + $n * $row_h + 34;
    my @rows;
    for my $i (0 .. $n-1) {
        my $mo = $months[$i];
        my $y = 4 + $i * $row_h;
        my @segs = ('Até 2','Até 5','Até 7','Acima 7');
        my @colors = ('var(--good)','var(--warning)','var(--serious)','var(--critical)');
        my $x = 60;
        my $rects = '';
        for my $si (0..3) {
            my $w = sprintf("%.1f", $mo->{$segs[$si]} * 5);
            $rects .= qq{<rect x="$x" y="$y" width="$w" height="20" fill="$colors[$si]"></rect>};
            $x += $w;
        }
        my $tip = sprintf("%s — até 2: %s%% · até 5: %s%% · até 7: %s%% · acima de 7: %s%%",
            $mo->{name}, fmt_pct1($mo->{'Até 2'}), fmt_pct1($mo->{'Até 5'}), fmt_pct1($mo->{'Até 7'}), fmt_pct1($mo->{'Acima 7'}));
        push @rows, qq{<g class="bar-row" data-tip="$tip"><text x="0" y="} . ($y+16) . qq{" class="row-label">$mo->{name}</text>$rects</g>};
    }
    my $axis_y = 4 + $n * $row_h + 4;
    my $svg = qq{<svg viewBox="0 0 640 $h" role="img" aria-label="Distribuicao percentual de pedidos por faixa de prazo">};
    $svg .= join('', @rows);
    $svg .= qq{<line x1="60" y1="$axis_y" x2="560" y2="$axis_y" stroke="var(--baseline)" stroke-width="1"></line>};
    my $label_y = $axis_y + 14;
    $svg .= qq{<text x="60" y="$label_y" class="axis-label">0%</text><text x="308" y="$label_y" class="axis-label">50%</text><text x="552" y="$label_y" class="axis-label">100%</text>};
    $svg .= qq{</svg>};
    return $svg;
}

# ---------------------------------------------------------------------------
# SVG: chart 2 - tendencia acima de 7 dias
# ---------------------------------------------------------------------------
sub trend_chart_svg {
    my (@months) = @_;
    my $n = scalar(@months);
    my ($x0, $x1) = (60, 540);
    my @pts;
    for my $i (0 .. $n-1) {
        my $x = $n > 1 ? $x0 + ($x1-$x0) * $i / ($n-1) : ($x0+$x1)/2;
        my $val = $months[$i]{'Acima 7'};
        my $y = 220 - $val * 2.0;
        push @pts, { x => sprintf("%.1f", $x), y => sprintf("%.1f", $y), val => $val, name => $months[$i]{name} };
    }
    my $path_line = join(' L ', map { "$_->{x},$_->{y}" } @pts);
    my $path_area = "M $pts[0]{x},220 L " . $path_line . " L $pts[-1]{x},220 Z";
    my $svg = qq{<svg viewBox="0 0 560 250" role="img" aria-label="Tendencia de pedidos acima de 7 dias">};
    $svg .= qq{<line x1="60" y1="20" x2="60" y2="220" stroke="var(--grid)" stroke-width="1"></line>};
    $svg .= qq{<line x1="60" y1="220" x2="540" y2="220" stroke="var(--baseline)" stroke-width="1"></line>};
    $svg .= qq{<line x1="60" y1="120" x2="540" y2="120" stroke="var(--grid)" stroke-width="1"></line>};
    $svg .= qq{<line x1="60" y1="20" x2="540" y2="20" stroke="var(--grid)" stroke-width="1"></line>};
    $svg .= qq{<text x="4" y="223" class="axis-label">0%</text><text x="4" y="123" class="axis-label">50%</text><text x="4" y="24" class="axis-label">100%</text>};
    $svg .= qq{<path d="$path_area" fill="var(--critical)" opacity="0.10"></path>};
    $svg .= qq{<path d="M $path_line" fill="none" stroke="var(--critical)" stroke-width="2"></path>};
    for my $i (0 .. $#pts) {
        my $p = $pts[$i];
        my $r = ($i == $#pts) ? 5 : 4;
        my $ring = ($i == $#pts) ? qq{ stroke="var(--surface)" stroke-width="2"} : '';
        my $tip = sprintf("%s — %s%% acima de 7 dias", $p->{name}, fmt_pct1($p->{val}));
        $svg .= qq{<circle cx="$p->{x}" cy="$p->{y}" r="$r" fill="var(--critical)"$ring data-tip="$tip"></circle>};
    }
    for my $p (@pts) {
        $svg .= qq{<text x="$p->{x}" y="240" class="axis-label" text-anchor="middle">$p->{name}</text>};
    }
    $svg .= qq{</svg>};
    return $svg;
}

# ---------------------------------------------------------------------------
# SVG: chart 3 - status diario mensal (verde/vermelho)
# ---------------------------------------------------------------------------
sub status_monthly_svg {
    my (@keys) = @_;
    my $n = scalar(@keys);
    my $row_h = 42;
    my $h = 10 + $n * $row_h;
    my @rows;
    for my $i (0 .. $n-1) {
        my $key = $keys[$i];
        my $rows_ref = $by_month{$key}{rows};
        my $total = scalar(@$rows_ref);
        my $g = scalar(grep { $_->{g} } @$rows_ref);
        my $pg = 100 * $g / $total;
        my $pv = 100 - $pg;
        my $y = 4 + $i * $row_h;
        my $wg = sprintf("%.1f", $pg * 3.0);
        my $wv = sprintf("%.1f", $pv * 3.0);
        my $xv = sprintf("%.1f", 55 + $pg * 3.0);
        my $tip = sprintf("%s — %s%% verde (%d dias) · %s%% vermelho (%d dias)",
            $by_month{$key}{name}, fmt_pct1($pg), $g, fmt_pct1($pv), $total - $g);
        push @rows, qq{<g class="bar-row" data-tip="$tip"><text x="0" y="} . ($y+16) . qq{" class="row-label">$by_month{$key}{name}</text><rect x="55" y="$y" width="$wg" height="20" fill="var(--good)"></rect><rect x="$xv" y="$y" width="$wv" height="20" fill="var(--critical)"></rect></g>};
    }
    my $svg = qq{<svg viewBox="0 0 380 $h" role="img" aria-label="Status diario verde ou vermelho por mes">};
    $svg .= join('', @rows);
    $svg .= qq{</svg>};
    return $svg;
}

# ---------------------------------------------------------------------------
# HEATMAP - calendario diario
# ---------------------------------------------------------------------------
sub heatmap_html {
    my (@keys) = @_;
    my $html = '';
    for my $key (@keys) {
        my $name = $by_month{$key}{name};
        my $mnum = $by_month{$key}{mnum};
        my $dots = '';
        for my $r (@{ $by_month{$key}{rows} }) {
            my $cls = $r->{g} ? 'g' : 'v';
            my $label = $r->{g} ? 'verde' : 'vermelho';
            my $mm = sprintf("%02d", $mnum);
            $dots .= qq{<span class="dot-day $cls" title="$r->{d}/$mm — $label"></span>};
        }
        $html .= qq{<div class="heatmap-month"><span class="m-label">$name</span><div class="heatmap-dots">$dots</div></div>};
    }
    return $html;
}

# ---------------------------------------------------------------------------
# LOGISTICA: graficos e blocos da secao de transporte
# ---------------------------------------------------------------------------
sub _cpp { my ($m)=@_; my $p = $pecas_mes{$m}//0; return $p > 0 ? ($gasto_mes{$m}//0)/$p : 0; }

# Grafico 1: evolucao indexada (base 100 no primeiro mes da janela).
# Se as duas linhas andam juntas, gasto e volume sao proporcionais.
sub log_chart_evol_svg {
    my (@ms) = @_;
    my $n = scalar(@ms);
    return qq{<svg viewBox="0 0 560 70" role="img"><text x="8" y="40" class="axis-label">Sem dados suficientes de transporte.</text></svg>} if $n < 2;
    my ($x0, $x1) = (58, 528);
    my ($ytop, $ybot) = (26, 196);
    my $bp = $pecas_mes{$ms[0]} || 1;
    my $bg = $gasto_mes{$ms[0]} || 1;
    my (@ip, @ig);
    for my $m (@ms) {
        push @ip, 100 * (($pecas_mes{$m} // 0) / $bp);
        push @ig, 100 * (($gasto_mes{$m} // 0) / $bg);
    }
    my $max = 100;
    for my $v (@ip, @ig) { $max = $v if $v > $max }
    $max *= 1.18;
    my $px = sub { my $i = shift; $n > 1 ? $x0 + ($x1-$x0)*$i/($n-1) : ($x0+$x1)/2 };
    my $py = sub { my $v = shift; sprintf("%.1f", $ybot - ($ybot-$ytop) * $v / $max) };

    my $svg = qq{<svg viewBox="0 0 560 236" role="img" aria-label="Evolucao indexada de pecas expedidas e gasto com transporte">};
    $svg .= qq{<line x1="$x0" y1="$ytop" x2="$x0" y2="$ybot" stroke="var(--grid)" stroke-width="1"></line>};
    $svg .= qq{<line x1="$x0" y1="$ybot" x2="$x1" y2="$ybot" stroke="var(--baseline)" stroke-width="1"></line>};
    # linha de referencia = 100 (base)
    my $y100 = $py->(100);
    $svg .= qq{<line x1="$x0" y1="$y100" x2="$x1" y2="$y100" stroke="var(--baseline)" stroke-width="1" stroke-dasharray="4 4"></line>};
    $svg .= qq{<text x="6" y="} . ($y100 + 4) . qq{" class="axis-label">100</text>};
    for my $ser ([\@ip, 'var(--accent)', 'pecas'], [\@ig, 'var(--serious)', 'gasto']) {
        my ($vals, $cor, $nome) = @$ser;
        my @pts = map { $px->($_) . "," . $py->($vals->[$_]) } (0 .. $n-1);
        $svg .= qq{<path d="M } . join(' L ', @pts) . qq{" fill="none" stroke="$cor" stroke-width="2.5"></path>};
        for my $i (0 .. $n-1) {
            my ($cx, $cy) = ($px->($i), $py->($vals->[$i]));
            my $real = $nome eq 'pecas'
                ? fmt_int($pecas_mes{$ms[$i]} // 0) . " pecas"
                : "R\$ " . fmt_brl($gasto_mes{$ms[$i]} // 0);
            my $tip = sprintf("%s — %s (indice %d)", $MESES_PT[$ms[$i]], $real, $vals->[$i] + 0.5);
            my $r = ($i == $n-1) ? 5 : 4;
            $svg .= qq{<circle cx="$cx" cy="$cy" r="$r" fill="$cor" stroke="var(--surface)" stroke-width="1.5" data-tip="$tip"></circle>};
        }
        # rotulo do indice no ultimo ponto
        my $lx = $px->($n-1) + 8;
        my $ly = $py->($vals->[$n-1]) + 4;
        my $lbl = sprintf("%d", $vals->[$n-1] + 0.5);
        $svg .= qq{<text x="$lx" y="$ly" class="axis-label" fill="$cor" style="font-weight:700">$lbl</text>};
    }
    for my $i (0 .. $n-1) {
        my $tx = $px->($i);
        $svg .= qq{<text x="$tx" y="220" class="axis-label" text-anchor="middle">$MESES_PT[$ms[$i]]</text>};
    }
    $svg .= qq{</svg>};
    return $svg;
}

# Grafico 2: custo por peca (R$), barras por mes.
sub log_chart_cpp_svg {
    my (@ms) = @_;
    my $n = scalar(@ms);
    return qq{<svg viewBox="0 0 560 70" role="img"><text x="8" y="40" class="axis-label">Sem dados.</text></svg>} if $n < 1;
    my ($x0, $x1) = (58, 528);
    my ($ytop, $ybot) = (26, 178);
    my $max = 0;
    for my $m (@ms) { my $c = _cpp($m); $max = $c if $c > $max }
    $max = $max > 0 ? $max * 1.3 : 1;
    my $soma = 0; $soma += _cpp($_) for @ms;
    my $media = $n ? $soma / $n : 0;
    my $largura = ($x1 - $x0) / $n;
    my $bw = $largura * 0.5;

    my $svg = qq{<svg viewBox="0 0 560 214" role="img" aria-label="Custo de transporte por peca expedida">};
    $svg .= qq{<line x1="$x0" y1="$ybot" x2="$x1" y2="$ybot" stroke="var(--baseline)" stroke-width="1"></line>};
    my $ym = sprintf("%.1f", $ybot - ($ybot-$ytop) * $media / $max);
    $svg .= qq{<line x1="$x0" y1="$ym" x2="$x1" y2="$ym" stroke="var(--muted)" stroke-width="1" stroke-dasharray="5 4"></line>};
    $svg .= qq{<text x="6" y="} . ($ym - 5) . qq{" class="axis-label">media R\$ } . fmt_brl3($media) . qq{</text>};
    for my $i (0 .. $n-1) {
        my $m = $ms[$i];
        my $c = _cpp($m);
        my $cx = $x0 + $largura * ($i + 0.5);
        my $bx = sprintf("%.1f", $cx - $bw/2);
        my $by = sprintf("%.1f", $ybot - ($ybot-$ytop) * $c / $max);
        my $bh = sprintf("%.1f", $ybot - $by);
        my $tip = sprintf("%s — R\$ %s por peca (R\$ %s / %s pecas)",
            $MESES_PT[$m], fmt_brl3($c), fmt_brl($gasto_mes{$m}//0), fmt_int($pecas_mes{$m}//0));
        $svg .= qq{<g class="bar-row" data-tip="$tip"><rect x="$bx" y="$by" width="} . sprintf("%.1f",$bw) . qq{" height="$bh" rx="3" fill="var(--accent)"></rect></g>};
        my $vy = $by - 7;
        $svg .= qq{<text x="$cx" y="$vy" class="axis-label" text-anchor="middle" style="font-weight:700">R\$ } . fmt_brl3($c) . qq{</text>};
        $svg .= qq{<text x="$cx" y="198" class="axis-label" text-anchor="middle">$MESES_PT[$m]</text>};
    }
    $svg .= qq{</svg>};
    return $svg;
}

# Tiles de KPI do mes mais recente da janela de logistica
sub log_kpis_html {
    my (@ms) = @_;
    return qq{<div class="card kpi"><div class="label">Transporte</div><div class="value">&#8212;</div><div class="sub">aba "$transp_nome" sem dados</div></div>} unless @ms;
    my $m = $ms[-1];
    my $ant = (scalar(@ms) >= 2) ? $ms[-2] : undef;
    my $g = $gasto_mes{$m} // 0;
    my $p = $pecas_mes{$m} // 0;
    my $v = $vol_mes{$m} // 0;
    my $cpp = _cpp($m);
    my $delta = '';
    if (defined $ant) {
        my $c0 = _cpp($ant);
        if ($c0 > 0) {
            my $var = 100 * ($cpp - $c0) / $c0;
            $delta = $var <= 0
                ? qq{<div class="delta good">&#9660; } . fmt_pct1(-$var) . qq{% vs $MESES_PT[$ant]</div>}
                : qq{<div class="delta bad">&#9650; } . fmt_pct1($var) . qq{% vs $MESES_PT[$ant]</div>};
        }
    }
    my $h = '';
    $h .= qq{<div class="card kpi"><div class="label">Gasto com transporte ($MESES_PT[$m])</div><div class="value">R\$ } . fmt_brl($g) . qq{</div><div class="sub">aba "$transp_nome"</div></div>};
    $h .= qq{<div class="card kpi"><div class="label">Custo por peça ($MESES_PT[$m])</div><div class="value">R\$ } . fmt_brl3($cpp) . qq{</div><div class="sub">} . fmt_int($p) . qq{ peças faturadas</div>$delta</div>};
    $h .= qq{<div class="card kpi"><div class="label">Volumes expedidos ($MESES_PT[$m])</div><div class="value">} . fmt_int($v) . qq{</div><div class="sub">} . ($v ? sprintf("%.0f", $p/$v) : '—') . qq{ peças por volume</div></div>};
    $h .= qq{<div class="card kpi"><div class="label">Custo por volume ($MESES_PT[$m])</div><div class="value">R\$ } . fmt_brl($v ? $g/$v : 0) . qq{</div><div class="sub">média do mês</div></div>};
    return $h;
}

# Tabela consolidada da janela
sub log_table_html {
    my (@ms) = @_;
    return '<p class="evidencia">Sem dados de transporte na planilha.</p>' unless @ms;
    my $h = qq{<table><thead><tr><th>Mês</th><th>Peças expedidas</th><th>Gasto transporte</th><th>Custo/peça</th><th>Volumes</th><th>Custo/volume</th><th>Peças/volume</th></tr></thead><tbody>};
    for my $m (@ms) {
        my $p = $pecas_mes{$m} // 0;
        my $g = $gasto_mes{$m} // 0;
        my $v = $vol_mes{$m} // 0;
        $h .= qq{<tr><td><b>$MESES_PT[$m]</b></td><td>} . fmt_int($p) . qq{</td><td>R\$ } . fmt_brl($g) . qq{</td><td><b>R\$ } . fmt_brl3(_cpp($m)) . qq{</b></td><td>} . fmt_int($v) . qq{</td><td>R\$ } . fmt_brl($v ? $g/$v : 0) . qq{</td><td>} . ($v ? sprintf("%.0f", $p/$v) : '—') . qq{</td></tr>};
    }
    $h .= qq{</tbody></table>};
    return $h;
}

# Achados calculados a partir dos numeros da janela
sub log_achados_html {
    my (@ms) = @_;
    return '' unless scalar(@ms) >= 2;
    my ($ini, $fim) = ($ms[0], $ms[-1]);
    my $dp = ($pecas_mes{$ini} // 0) > 0 ? 100 * (($pecas_mes{$fim}//0) - $pecas_mes{$ini}) / $pecas_mes{$ini} : 0;
    my $dg = ($gasto_mes{$ini} // 0) > 0 ? 100 * (($gasto_mes{$fim}//0) - $gasto_mes{$ini}) / $gasto_mes{$ini} : 0;
    my $h = '';

    # 1) proporcionalidade
    my $veredito;
    my $folga = abs($dg - $dp);
    if ($folga < 10) {
        $veredito = qq{<b>Sim, acompanham.</b> No período, peças variaram } . fmt_pct1($dp) . qq{% e o gasto } . fmt_pct1($dg) . qq{% — diferença de } . fmt_pct1($folga) . qq{ p.p., dentro do esperado.};
    } else {
        my $pior = $dg > $dp ? "o gasto subiu mais rápido que o volume" : "o gasto cresceu menos que o volume (ganho de eficiência)";
        $veredito = qq{<b>Não.</b> De $MESES_PT[$ini] a $MESES_PT[$fim] as peças variaram } . fmt_pct1($dp) . qq{% enquanto o gasto variou } . fmt_pct1($dg) . qq{% — } . $pior . qq{, uma diferença de } . fmt_pct1($folga) . qq{ p.p.};
    }
    $h .= qq{<div class="card causa"><span class="stripe" style="background:var(--serious)"></span><div class="area">Gasto x volume são proporcionais?</div><p class="problema">$veredito</p><p class="evidencia">Abas $transp_nome (gasto) e Relatório (Qtd. Faturada)</p></div>};

    # 2) melhor e pior custo por peca
    my ($best, $worst) = ($ms[0], $ms[0]);
    for my $m (@ms) { $best = $m if _cpp($m) < _cpp($best); $worst = $m if _cpp($m) > _cpp($worst); }
    if ($best != $worst) {
        my $gap = _cpp($best) > 0 ? 100 * (_cpp($worst) - _cpp($best)) / _cpp($best) : 0;
        my $econ = ($pecas_mes{$fim}//0) * (_cpp($fim) - _cpp($best));
        my $frase = $econ > 0
            ? qq{ Aplicando em $MESES_PT[$fim] o custo unitário de $MESES_PT[$best], o gasto do mês cairia cerca de <b>R\$ } . fmt_brl($econ) . qq{</b>.}
            : '';
        $h .= qq{<div class="card causa"><span class="stripe" style="background:var(--critical)"></span><div class="area">Amplitude do custo unitário</div><p class="problema">O custo por peça oscilou <b>} . fmt_pct1($gap) . qq{%</b> entre $MESES_PT[$best] (R\$ } . fmt_brl3(_cpp($best)) . qq{) e $MESES_PT[$worst] (R\$ } . fmt_brl3(_cpp($worst)) . qq{).$frase</p><p class="evidencia">Variação grande demais para ser só mix de destino — indica falta de política de frete</p></div>};
    }

    # 3) concentracao de transportadora
    if ($emp_mes{$fim}) {
        my @k = sort { $emp_mes{$fim}{$b} <=> $emp_mes{$fim}{$a} } keys %{$emp_mes{$fim}};
        my $tot = 0; $tot += $emp_mes{$fim}{$_} for @k;
        if ($tot > 0) {
            my $pct = 100 * $emp_mes{$fim}{$k[0]} / $tot;
            my $lista = join(", ", map { "$_ (" . sprintf("%.0f", 100*$emp_mes{$fim}{$_}/$tot) . "%)" } @k[0 .. ($#k > 2 ? 2 : $#k)]);
            my $risco = $pct >= 90
                ? qq{Dependência praticamente total de um único fornecedor: sem alternativa cotada, não há como saber se o preço está competitivo nem plano B se faltar veículo.}
                : qq{Distribuição entre fornecedores permite comparação de preço.};
            $h .= qq{<div class="card causa"><span class="stripe" style="background:} . ($pct >= 90 ? 'var(--critical)' : 'var(--good)') . qq{"></span><div class="area">Concentração de fornecedor</div><p class="problema">Em $MESES_PT[$fim], <b>} . fmt_pct1($pct) . qq{%</b> dos lançamentos foram de <b>$k[0]</b>. $risco</p><p class="evidencia">Mix: $lista</p></div>};
        }
    }

    # 4) qualidade do dado: valores gravados como texto
    my $txt_total = 0; $txt_total += ($txt_mes{$_} // 0) for @ms;
    if ($txt_total > 0) {
        my @det = map { "$MESES_PT[$_]: " . $txt_mes{$_} } grep { ($txt_mes{$_}//0) > 0 } @ms;
        $h .= qq{<div class="card causa"><span class="stripe" style="background:var(--muted)"></span><div class="area">Qualidade do dado de frete</div><p class="problema"><b>$txt_total lançamento(s)</b> estão com o valor gravado como texto (ex: <code>218;89</code> com ponto-e-vírgula no lugar da vírgula). Esses valores <b>não entram no total do mês</b>, então o gasto real é maior que o exibido.</p><p class="evidencia">} . join(" · ", @det) . qq{ — corrigir na origem e aplicar validação de dados na coluna Valor</p></div>};
    }
    return $h;
}

my $log_periodo = @log_months
    ? ($MESES_PT[$log_months[0]] . (scalar(@log_months) > 1 ? "\x{2013}" . $MESES_PT[$log_months[-1]] : ''))
    : '—';

# ---------------------------------------------------------------------------
# BUILD OUTPUT
# ---------------------------------------------------------------------------
my $chart1 = bucket_chart_svg(@active_months);
my $chart2 = trend_chart_svg(@active_months);
my $chart3 = status_monthly_svg(@month_keys);
my $heatmap = heatmap_html(@month_keys);

my $last_month_name = $last->{name};
my $last_month_abbr  = lc($MESES_ABR[ (grep { $MESES_PT[$_] eq $last_month_name } (1..12))[0] ]);

my %tokens = (
    '{{PERIODO}}'              => $periodo_label,
    '{{TOTAL_DIAS}}'           => $total_dias,
    '{{META_ATE2}}'            => sprintf("%d", $meta_ate2),
    '{{META_ATE5}}'            => sprintf("%d", $meta_ate5),
    '{{META_ATE7}}'            => sprintf("%d", $meta_ate7),
    '{{META_ACIMA7}}'          => sprintf("%d", $meta_acima7),
    '{{KPI_ATE2_LAST}}'        => fmt_pct1($kpi_ate2_last),
    '{{KPI_ATE5_LAST}}'        => fmt_pct1($kpi_ate5_last),
    '{{KPI_ATE7_LAST}}'        => fmt_pct1($kpi_ate7_last),
    '{{KPI_ACIMA7_LAST}}'      => fmt_pct1($kpi_acima7_last),
    '{{KPI_LAST_MONTH}}'       => $last_month_abbr,
    '{{KPI_ATE2_GAP_HTML}}'    => $kpi_ate2_gap_html,
    '{{KPI_ATE5_GAP_HTML}}'    => $kpi_ate5_gap_html,
    '{{KPI_ATE7_GAP_HTML}}'    => $kpi_ate7_gap_html,
    '{{KPI_ACIMA7_GAP_HTML}}'  => $kpi_acima7_gap_html,
    '{{PCT_VERMELHO_TOTAL}}'   => fmt_pct1($pct_vermelho_total),
    '{{N_VERMELHO_TOTAL}}'     => ($total_dias - $total_verde),
    '{{N_DIAS_TOTAL}}'         => $total_dias,
    '{{CHART_BUCKETS_SVG}}'    => $chart1,
    '{{CHART_TREND_SVG}}'      => $chart2,
    '{{CHART_STATUS_SVG}}'     => $chart3,
    '{{HEATMAP_HTML}}'         => $heatmap,
    '{{GENERATED_AT}}'         => sprintf("%04d-%02d-%02d %02d:%02d", $today_y, $today_m, $today_d, $today_h, $today_min),
    '{{LAST_MONTH_NAME}}'      => $last_month_name,
    '{{LOG_ABA}}'              => $transp_nome,
    '{{LOG_PERIODO}}'          => $log_periodo,
    '{{LOG_KPIS_HTML}}'        => log_kpis_html(@log_months),
    '{{LOG_CHART_EVOL_SVG}}'   => log_chart_evol_svg(@log_months),
    '{{LOG_CHART_CPP_SVG}}'    => log_chart_cpp_svg(@log_months),
    '{{LOG_TABLE_HTML}}'       => log_table_html(@log_months),
    '{{LOG_ACHADOS_HTML}}'     => log_achados_html(@log_months),
);

open(my $tf, '<:encoding(UTF-8)', $TEMPLATE_FILE) or die "$TEMPLATE_FILE: $!";
local $/;
my $tpl = <$tf>;
close $tf;

for my $k (keys %tokens) {
    my $v = $tokens{$k};
    $tpl =~ s/\Q$k\E/$v/g;
}

open(my $of, '>:encoding(UTF-8)', $OUTPUT_FILE) or die "$OUTPUT_FILE: $!";
print $of $tpl;
close $of;

print "OK — gerado $OUTPUT_FILE\n";
print "Meses ativos: ", join(", ", map { $_->{name} } @active_months), "\n";
print "Ultimo mes: $last_month_name  Ate2=$kpi_ate2_last%  Acima7=$kpi_acima7_last%\n";
print "Total dias=$total_dias  vermelho=$pct_vermelho_total%\n";