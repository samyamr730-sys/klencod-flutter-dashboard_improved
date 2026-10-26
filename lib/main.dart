import 'dart:async';
import 'dart:convert';
import 'dart:isolate';
import 'dart:math' as math;
import 'dart:typed_data';
import 'package:excel/excel.dart' hide Border;
import 'package:file_picker/file_picker.dart';
import 'package:fl_chart/fl_chart.dart';
import 'package:flutter/material.dart';
import 'package:flutter_localizations/flutter_localizations.dart';

const bg = Color(0xFF0A1226);
const card = Color(0xFF111B33);
const card2 = Color(0xFF16223F);
const border = Color(0xFF1E2D52);
const muted = Color(0xFF8DA2C9);
const blue = Color(0xFF2F6BFF);
const green = Color(0xFF12C98A);
const purple = Color(0xFF8B5CF6);
const orange = Color(0xFFFF8A1F);
const pink = Color(0xFFFF4D9D);
const cyan = Color(0xFF22D3EE);
const palette = [blue, green, purple, orange, pink, cyan];

const int pageSize = 50;
const int maxRows = 100000;
const int maxLinePoints = 30;

void main() => runApp(const App());

// ───────────────────────── أدوات الأرقام ─────────────────────────

String _normDigits(String s) {
  final b = StringBuffer();
  for (final r in s.runes) {
    if (r >= 0x660 && r <= 0x669) {
      b.writeCharCode(r - 0x660 + 48);
    } else if (r >= 0x6F0 && r <= 0x6F9) {
      b.writeCharCode(r - 0x6F0 + 48);
    } else if (r == 0x66B) {
      b.write('.');
    } else if (r == 0x66C) {
      // فاصل آلاف عربي: نتجاهله
    } else {
      b.writeCharCode(r);
    }
  }
  return b.toString();
}

final RegExp _noise = RegExp(r'[\s%$€£\u00A0]|EGP|SAR|USD|ج\.م|جنيه', caseSensitive: false);
final RegExp _decComma = RegExp(r'^[-+]?\d+,\d{1,2}$');

double? toNum(String s) {
  var t = s.trim();
  if (t.isEmpty) return null;
  t = _normDigits(t).replaceAll(_noise, '');
  if (_decComma.hasMatch(t)) {
    t = t.replaceAll(',', '.');
  } else {
    t = t.replaceAll(',', '');
  }
  final v = double.tryParse(t);
  return (v != null && v.isFinite) ? v : null;
}

String compact(double v) {
  final a = v.abs();
  String f(double x) => x == x.roundToDouble() ? x.toStringAsFixed(0) : x.toStringAsFixed(1);
  if (a >= 1e9) return '${f(v / 1e9)}B';
  if (a >= 1e6) return '${f(v / 1e6)}M';
  if (a >= 1e3) return '${f(v / 1e3)}K';
  return v == v.roundToDouble() ? v.toStringAsFixed(0) : v.toStringAsFixed(2);
}

/// مدى المحور الرأسي بدون قيم صفرية/متساوية تسبب مشاكل في الرسم.
List<double> yRange(List<double> v, {required bool zero}) {
  var lo = v.reduce(math.min);
  var hi = v.reduce(math.max);
  if (zero) {
    lo = math.min(lo, 0);
    hi = math.max(hi, 0);
  }
  if (hi - lo < 1e-9) {
    hi = lo + 1;
  } else {
    final pad = (hi - lo) * 0.08;
    hi += pad;
    if (!zero) lo -= pad;
  }
  return [lo, hi];
}

// ───────────────────────── نموذج البيانات ─────────────────────────

bool _idLike(String h, List<String> vals) {
  final t = h.trim().toLowerCase();
  if (RegExp(r'^(id|#|م|no\.?|sr\.?|serial|index|رقم|الرقم|كود|تسلسل|مسلسل)$').hasMatch(t)) return true;
  if (RegExp(r'(^|[\s_])id($|[\s_])').hasMatch(t)) return true;
  if (t.length <= 12 && RegExp(r'year|سنة|السنة|عام|العام').hasMatch(t)) return true;
  final ns = vals.map(toNum).whereType<double>().toList();
  if (ns.length >= 3 && ns.every((x) => x == x.roundToDouble()) && ns.toSet().length == ns.length) {
    var seq = true;
    for (var i = 1; i < ns.length; i++) {
      if (ns[i] - ns[i - 1] != 1) {
        seq = false;
        break;
      }
    }
    if (seq) return true;
  }
  return false;
}

class Tbl {
  final List<String> head;
  final List<List<String>> rows;
  final List<int> numCols;
  final List<int> catCols;
  final int? defCat;
  final int? defVal;
  List<String>? _hay;
  Tbl(this.head, this.rows, this.numCols, this.catCols, this.defCat, this.defVal);

  List<String> get hay => _hay ??= rows.map((r) => r.join('\u0001').toLowerCase()).toList();

  static Tbl? from(List<List<String>> input) {
    var all = input.where((r) => r.any((c) => c.trim().isNotEmpty)).toList();
    if (all.length < 2) return null;

    // اكتشاف صف العناوين (يتجاوز عنوان التقرير في أعلى الشيت)
    final probe = all.take(10).toList();
    final counts = probe.map((r) => r.where((c) => c.trim().isNotEmpty).length).toList();
    final mxc = counts.reduce(math.max);
    final need = (mxc * 0.6).ceil();
    var h = 0;
    for (var i = 0; i < probe.length; i++) {
      final cells = probe[i].map((c) => c.trim()).where((c) => c.isNotEmpty).toList();
      final nf = cells.where((c) => toNum(c) != null).length / cells.length;
      if (counts[i] >= need && nf < 0.5) {
        h = i;
        break;
      }
    }
    if (h >= all.length - 1) h = 0;
    if (h > 0) all = all.sublist(h);
    final raw = all;

    final width = raw.map((r) => r.length).reduce(math.max);
    final keep = <int>[];
    for (var c = 0; c < width; c++) {
      for (var r = 1; r < raw.length; r++) {
        if (c < raw[r].length && raw[r][c].trim().isNotEmpty) {
          keep.add(c);
          break;
        }
      }
    }
    if (keep.isEmpty) return null;
    final head = keep.map((c) {
      final t = c < raw[0].length ? raw[0][c].trim() : '';
      return t.isEmpty ? 'عمود ${c + 1}' : t;
    }).toList();
    final rows = <List<String>>[];
    for (var r = 1; r < raw.length; r++) {
      final row = keep.map((c) => c < raw[r].length ? raw[r][c].trim() : '').toList();
      if (row.any((x) => x.isNotEmpty)) rows.add(row);
    }
    if (rows.isEmpty) return null;

    final nums = <int>[];
    final cats = <int>[];
    final idCols = <int>{};
    int? defCat;
    for (var i = 0; i < head.length; i++) {
      final vals = rows.map((r) => r[i]).where((x) => x.isNotEmpty).toList();
      if (vals.isEmpty) continue;
      final n = vals.where((x) => toNum(x) != null).length;
      if (n / vals.length >= 0.7) {
        nums.add(i);
        if (_idLike(head[i], vals)) idCols.add(i);
        // أعمدة رقمية قليلة القيم (سنة، شهر، مجموعة) تصلح كتصنيف أيضاً
        if (vals.toSet().length <= 15 || idCols.contains(i)) cats.add(i);
      } else {
        cats.add(i);
        final u = vals.toSet().length;
        if (defCat == null && u >= 2 && u <= rows.length * 0.9) defCat = i;
      }
    }
    cats.sort();
    defCat ??= cats.isEmpty ? null : cats.first;
    int? defVal;
    for (final i in nums) {
      if (!idCols.contains(i) && i != defCat) {
        defVal = i;
        break;
      }
    }
    defVal ??= nums.isEmpty ? null : nums.first;
    return Tbl(head, rows, nums, cats, defCat, defVal);
  }
}

class Agg {
  final double sum, avg, mx, mn, med;
  final int groups;
  final List<MapEntry<String, double>> ordered, top;
  Agg(this.sum, this.avg, this.mx, this.mn, this.med, this.groups, this.ordered, this.top);
}

Agg computeAgg(Tbl s, int? vi, int? ci, int agg) {
  final vals = <double>[];
  final groups = <String, List<double>>{};
  if (vi != null) {
    for (var k = 0; k < s.rows.length; k++) {
      final x = toNum(s.rows[k][vi]);
      if (x == null) continue;
      vals.add(x);
      final String key;
      if (ci == null) {
        key = '#${k + 1}';
      } else {
        final c = s.rows[k][ci];
        key = c.isEmpty ? 'بدون تصنيف' : c;
      }
      (groups[key] ??= []).add(x);
    }
  }
  final sum = vals.fold<double>(0, (a, b) => a + b);
  final avg = vals.isEmpty ? 0.0 : sum / vals.length;
  var mx = 0.0, mn = 0.0, med = 0.0;
  if (vals.isNotEmpty) {
    final sorted = [...vals]..sort();
    mx = sorted.last;
    mn = sorted.first;
    final m = sorted.length ~/ 2;
    med = sorted.length.isOdd ? sorted[m] : (sorted[m - 1] + sorted[m]) / 2;
  }
  final agged = groups.entries.map((e) {
    final t = e.value.reduce((a, b) => a + b);
    return MapEntry(e.key, agg == 0 ? t : t / e.value.length);
  }).toList();
  final top = [...agged]..sort((a, b) => b.value.compareTo(a.value));
  return Agg(sum, avg, mx, mn, med, agged.length, agged.take(maxLinePoints).toList(), top);
}

// ───────────────────────── قراءة الملفات (في خيط خلفي) ─────────────────────────

List<List<String>> parseCsv(String text) {
  var t = text;
  if (t.startsWith('\uFEFF')) t = t.substring(1);
  final nl = t.indexOf('\n');
  final first = nl < 0 ? t : t.substring(0, nl);
  var d = ',';
  var best = first.split(',').length;
  for (final c in [';', '\t']) {
    final n = first.split(c).length;
    if (n > best) {
      best = n;
      d = c;
    }
  }
  final rows = <List<String>>[];
  var row = <String>[];
  final sb = StringBuffer();
  var inQ = false;
  for (var i = 0; i < t.length; i++) {
    final ch = t[i];
    if (inQ) {
      if (ch == '"') {
        if (i + 1 < t.length && t[i + 1] == '"') {
          sb.write('"');
          i++;
        } else {
          inQ = false;
        }
      } else {
        sb.write(ch);
      }
    } else if (ch == '"') {
      inQ = true;
    } else if (ch == d) {
      row.add(sb.toString());
      sb.clear();
    } else if (ch == '\n' || ch == '\r') {
      if (ch == '\r' && i + 1 < t.length && t[i + 1] == '\n') i++;
      row.add(sb.toString());
      sb.clear();
      rows.add(row);
      row = <String>[];
    } else {
      sb.write(ch);
    }
  }
  if (sb.isNotEmpty || row.isNotEmpty) {
    row.add(sb.toString());
    rows.add(row);
  }
  return rows;
}

class Loaded {
  final Map<String, Tbl> sheets;
  final bool formulas;
  final bool cut;
  Loaded(this.sheets, this.formulas, this.cut);
}

Loaded loadFile(Uint8List bytes, bool csv) {
  final m = <String, Tbl>{};
  var formulas = false;
  var cut = false;
  if (csv) {
    var raw = parseCsv(utf8.decode(bytes, allowMalformed: true));
    if (raw.length > maxRows + 1) {
      raw = raw.sublist(0, maxRows + 1);
      cut = true;
    }
    final s = Tbl.from(raw);
    if (s != null) m['CSV'] = s;
  } else {
    final ex = Excel.decodeBytes(bytes);
    ex.tables.forEach((n, t) {
      final src = t.rows;
      final lim = src.length > maxRows + 1 ? maxRows + 1 : src.length;
      if (src.length > lim) cut = true;
      final raw = <List<String>>[];
      for (var i = 0; i < lim; i++) {
        raw.add(src[i].map((c) {
          final v = c?.value;
          if (v == null) return '';
          if (v.runtimeType.toString() == 'FormulaCellValue') {
            formulas = true;
            return '';
          }
          return v.toString();
        }).toList());
      }
      final s = Tbl.from(raw);
      if (s != null) m[n] = s;
    });
  }
  return Loaded(m, formulas, cut);
}

// دالة على مستوى الملف (بدون التقاط this) لتشغيلها في Isolate
Future<Loaded> loadInBackground(Uint8List bytes, bool csv) => Isolate.run(() => loadFile(bytes, csv));

// ───────────────────────── التطبيق ─────────────────────────

class App extends StatelessWidget {
  const App({super.key});
  @override
  Widget build(BuildContext c) => MaterialApp(
        title: 'داش بورد الإكسيل',
        debugShowCheckedModeBanner: false,
        locale: const Locale('ar'),
        supportedLocales: const [Locale('ar'), Locale('en')],
        localizationsDelegates: const [
          GlobalMaterialLocalizations.delegate,
          GlobalWidgetsLocalizations.delegate,
          GlobalCupertinoLocalizations.delegate,
        ],
        theme: ThemeData.dark(useMaterial3: true).copyWith(
          scaffoldBackgroundColor: bg,
          colorScheme: const ColorScheme.dark(primary: blue, surface: card),
        ),
        home: const Home(),
      );
}

class Home extends StatefulWidget {
  const Home({super.key});
  @override
  State<Home> createState() => _HomeState();
}

class _HomeState extends State<Home> {
  Map<String, Tbl> sheets = {};
  String? sheetName;
  String fileName = '';
  bool loading = false;
  String? error;
  int? catIdx;
  int? valIdx;
  int agg = 0;
  int chartType = 0;
  String q = '';
  int? sortCol;
  bool asc = true;
  int page = 0;

  final qc = TextEditingController();
  Timer? _deb;

  // كاش الحسابات
  Agg? _agg;
  String _aggKey = '';
  String _tKey = '';
  List<List<String>> _tRows = [];

  @override
  void dispose() {
    _deb?.cancel();
    qc.dispose();
    super.dispose();
  }

  void select(String n) {
    sheetName = n;
    final s = sheets[n]!;
    catIdx = s.defCat;
    valIdx = s.defVal;
    q = '';
    qc.clear();
    sortCol = null;
    page = 0;
  }

  void notify(String m) {
    if (!mounted) return;
    if (sheetName == null) {
      setState(() => error = m);
    } else {
      ScaffoldMessenger.of(context)
        ..hideCurrentSnackBar()
        ..showSnackBar(SnackBar(content: Text(m), behavior: SnackBarBehavior.floating));
    }
  }

  Future<void> pick() async {
    try {
      final r = await FilePicker.platform
          .pickFiles(type: FileType.custom, allowedExtensions: ['xlsx', 'csv'], withData: true);
      if (r == null) return;
      final f = r.files.single;
      final bytes = f.bytes;
      if (bytes == null) {
        notify('تعذر قراءة محتوى الملف');
        return;
      }
      setState(() {
        loading = true;
        error = null;
      });
      final isCsv = f.name.toLowerCase().endsWith('.csv');
      final res = await loadInBackground(bytes, isCsv);
      if (!mounted) return;
      if (res.sheets.isEmpty) {
        setState(() => loading = false);
        notify('الملف فارغ أو لا يحتوي بيانات');
        return;
      }
      setState(() {
        loading = false;
        error = null;
        sheets = res.sheets;
        fileName = f.name;
        select(res.sheets.keys.first);
      });
      final notes = <String>[];
      if (res.formulas) notes.add('بعض الخلايا صيغ لا يمكن حسابها هنا، احفظ الملف "كقيم" ثم أعد رفعه.');
      if (res.cut) notes.add('تم تحميل أول $maxRows صف فقط.');
      if (notes.isNotEmpty) notify(notes.join('\n'));
    } catch (e) {
      if (!mounted) return;
      setState(() => loading = false);
      notify('تعذر قراءة الملف، تأكد أنه xlsx أو csv سليم.');
    }
  }

  void loadSample() {
    final rnd = math.Random(7);
    const months = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
    const prods = ['منتج أ', 'منتج ب', 'منتج ج'];
    final raw = <List<String>>[
      ['الشهر', 'المنتج', 'المبيعات', 'التكلفة']
    ];
    for (final m in months) {
      for (final p in prods) {
        final s = 800 + rnd.nextInt(1200);
        raw.add([m, p, '$s', '${(s * (0.5 + rnd.nextDouble() * 0.3)).round()}']);
      }
    }
    setState(() {
      sheets = {'عينة': Tbl.from(raw)!};
      fileName = 'بيانات تجريبية';
      error = null;
      select('عينة');
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: SafeArea(
        child: loading
            ? const Center(child: CircularProgressIndicator())
            : sheetName == null
                ? welcome()
                : dashboard(),
      ),
    );
  }

  Widget welcome() => Center(
        child: SingleChildScrollView(
          padding: const EdgeInsets.all(28),
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            Container(
              padding: const EdgeInsets.all(22),
              decoration: BoxDecoration(
                gradient: const LinearGradient(colors: [blue, purple]),
                borderRadius: BorderRadius.circular(26),
              ),
              child: const Icon(Icons.insert_chart_rounded, size: 56, color: Colors.white),
            ),
            const SizedBox(height: 22),
            const Text('داش بورد الإكسيل', style: TextStyle(fontSize: 26, fontWeight: FontWeight.bold)),
            const SizedBox(height: 8),
            const Text('ارفع ملف Excel أو CSV وهتظهر لك الإحصائيات والرسوم فوراً',
                textAlign: TextAlign.center, style: TextStyle(color: muted)),
            if (error != null) ...[
              const SizedBox(height: 14),
              Text(error!, style: const TextStyle(color: pink), textAlign: TextAlign.center),
            ],
            const SizedBox(height: 26),
            FilledButton.icon(
              style: FilledButton.styleFrom(
                  backgroundColor: blue, padding: const EdgeInsets.symmetric(horizontal: 28, vertical: 14)),
              onPressed: pick,
              icon: const Icon(Icons.upload_file),
              label: const Text('ارفع ملف'),
            ),
            const SizedBox(height: 10),
            TextButton.icon(
              onPressed: loadSample,
              icon: const Icon(Icons.auto_awesome_rounded),
              label: const Text('جرّب بيانات تجريبية'),
            ),
          ]),
        ),
      );

  Agg aggFor(Tbl s) {
    final key = '${identityHashCode(s)}|$catIdx|$valIdx|$agg';
    if (_agg == null || key != _aggKey) {
      _aggKey = key;
      _agg = computeAgg(s, valIdx, catIdx, agg);
    }
    return _agg!;
  }

  Widget dashboard() {
    final s = sheets[sheetName]!;
    final vi = valIdx;
    final ci = catIdx;
    final a = aggFor(s);
    final useLine = chartType == 1 && a.ordered.length >= 2;
    final scale = MediaQuery.textScalerOf(context).scale(1.0).clamp(1.0, 1.6).toDouble();

    return ListView(padding: const EdgeInsets.all(16), children: [
      Row(children: [
        Container(
          padding: const EdgeInsets.all(10),
          decoration: BoxDecoration(
              gradient: const LinearGradient(colors: [blue, purple]), borderRadius: BorderRadius.circular(14)),
          child: const Icon(Icons.insert_chart_rounded, color: Colors.white),
        ),
        const SizedBox(width: 12),
        Expanded(
          child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
            const Text('داش بورد الإكسيل', style: TextStyle(fontSize: 20, fontWeight: FontWeight.bold)),
            Text(fileName, maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(color: muted, fontSize: 12)),
          ]),
        ),
        IconButton.filledTonal(onPressed: pick, icon: const Icon(Icons.upload_file)),
      ]),
      const SizedBox(height: 18),
      Panel(
        title: 'الإعدادات',
        icon: Icons.tune_rounded,
        child: Column(children: [
          if (sheets.length > 1)
            Padding(
              padding: const EdgeInsets.only(bottom: 10),
              child: labeled(
                'الشيت',
                DropdownButtonHideUnderline(
                  child: DropdownButton<String>(
                    value: sheetName,
                    isExpanded: true,
                    dropdownColor: card2,
                    items: sheets.keys.map((k) => DropdownMenuItem(value: k, child: Text(k))).toList(),
                    onChanged: (v) => setState(() => select(v!)),
                  ),
                ),
              ),
            ),
          Row(children: [
            Expanded(child: labeled('عمود التصنيف', colDrop(s, s.catCols, ci, (v) => setState(() => catIdx = v)))),
            const SizedBox(width: 10),
            Expanded(child: labeled('عمود القيمة', colDrop(s, s.numCols, vi, (v) => setState(() => valIdx = v)))),
          ]),
          const SizedBox(height: 10),
          Row(children: [
            const Text('طريقة التجميع:', style: TextStyle(color: muted)),
            const SizedBox(width: 10),
            ChoiceChip(label: const Text('مجموع'), selected: agg == 0, onSelected: (_) => setState(() => agg = 0)),
            const SizedBox(width: 8),
            ChoiceChip(label: const Text('متوسط'), selected: agg == 1, onSelected: (_) => setState(() => agg = 1)),
          ]),
        ]),
      ),
      GridView(
        shrinkWrap: true,
        physics: const NeverScrollableScrollPhysics(),
        gridDelegate: SliverGridDelegateWithFixedCrossAxisCount(
          crossAxisCount: 2,
          mainAxisSpacing: 12,
          crossAxisSpacing: 12,
          mainAxisExtent: 112 * scale,
        ),
        children: [
          Kpi('الصفوف', '${s.rows.length}', Icons.table_rows_rounded, const [Color(0xFF1D4ED8), blue]),
          Kpi('المجموع', compact(a.sum), Icons.functions_rounded, const [Color(0xFF059669), green]),
          Kpi('المتوسط', compact(a.avg), Icons.show_chart_rounded, const [Color(0xFF6D28D9), purple]),
          Kpi('الوسيط', compact(a.med), Icons.align_vertical_center_rounded, const [Color(0xFF0E7490), cyan]),
          Kpi('الأعلى', compact(a.mx), Icons.trending_up_rounded, const [Color(0xFFEA580C), Color(0xFFFF9A3D)]),
          Kpi('الأدنى', compact(a.mn), Icons.trending_down_rounded, const [Color(0xFFBE185D), pink]),
        ],
      ),
      const SizedBox(height: 14),
      if (vi == null)
        const Panel(
          title: 'تنبيه',
          icon: Icons.info_outline,
          child: Text('لا توجد أعمدة رقمية في هذا الشيت لعرض الرسوم.', style: TextStyle(color: muted)),
        )
      else ...[
        Panel(
          title: 'نظرة عامة',
          icon: Icons.area_chart_rounded,
          trailing: Wrap(spacing: 6, children: [
            ChoiceChip(
                label: const Text('أعمدة'),
                visualDensity: VisualDensity.compact,
                selected: !useLine,
                onSelected: (_) => setState(() => chartType = 0)),
            ChoiceChip(
                label: const Text('خطي'),
                visualDensity: VisualDensity.compact,
                selected: useLine,
                onSelected: a.ordered.length >= 2 ? (_) => setState(() => chartType = 1) : null),
          ]),
          child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
            SizedBox(
              height: 250,
              child: useLine ? lineChart(a.ordered) : barChart(a.top.take(8).toList()),
            ),
            const SizedBox(height: 8),
            Text(
              useLine
                  ? (a.groups > maxLinePoints
                      ? 'أول $maxLinePoints فئة بترتيب الملف من ${a.groups}'
                      : '${a.groups} فئة بترتيب الملف')
                  : (a.groups > 8 ? 'أعلى 8 فئات من ${a.groups}' : '${a.groups} فئة'),
              style: const TextStyle(color: muted, fontSize: 11),
            ),
          ]),
        ),
        Panel(title: 'التوزيع', icon: Icons.donut_large_rounded, child: donut(a.top)),
        Panel(title: 'الأعلى قيمة', icon: Icons.leaderboard_rounded, child: ranked(a.top.take(6).toList())),
      ],
      Panel(title: 'البيانات', icon: Icons.table_chart_rounded, child: table(s)),
    ]);
  }

  Widget labeled(String l, Widget w) => Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
        Text(l, style: const TextStyle(color: muted, fontSize: 12)),
        const SizedBox(height: 4),
        Container(
          padding: const EdgeInsets.symmetric(horizontal: 12),
          decoration: BoxDecoration(color: card2, borderRadius: BorderRadius.circular(12), border: Border.all(color: border)),
          child: w,
        ),
      ]);

  Widget colDrop(Tbl s, List<int> opts, int? cur, ValueChanged<int?> on) {
    if (opts.isEmpty) {
      return const Padding(padding: EdgeInsets.symmetric(vertical: 14), child: Text('—'));
    }
    return DropdownButtonHideUnderline(
      child: DropdownButton<int>(
        value: opts.contains(cur) ? cur : null,
        isExpanded: true,
        dropdownColor: card2,
        hint: const Text('اختر'),
        items: opts.map((i) => DropdownMenuItem(value: i, child: Text(s.head[i], overflow: TextOverflow.ellipsis))).toList(),
        onChanged: on,
      ),
    );
  }

  Widget label(double v, List<MapEntry<String, double>> d) {
    final i = v.toInt();
    if (v != v.roundToDouble() || i < 0 || i >= d.length) return const SizedBox();
    final t = d[i].key;
    return Padding(
      padding: const EdgeInsets.only(top: 6),
      child: Text(t.length > 6 ? '${t.substring(0, 6)}..' : t, style: const TextStyle(color: muted, fontSize: 10)),
    );
  }

  FlTitlesData titles(List<MapEntry<String, double>> d, double interval, double yi) => FlTitlesData(
        topTitles: const AxisTitles(),
        rightTitles: const AxisTitles(),
        leftTitles: AxisTitles(
          sideTitles: SideTitles(
            showTitles: true,
            reservedSize: 44,
            interval: yi,
            getTitlesWidget: (v, m) => Text(compact(v), style: const TextStyle(color: muted, fontSize: 10)),
          ),
        ),
        bottomTitles: AxisTitles(
          sideTitles: SideTitles(
            showTitles: true,
            reservedSize: 30,
            interval: interval,
            getTitlesWidget: (v, m) => label(v, d),
          ),
        ),
      );

  FlGridData grid(double yi) => FlGridData(
        show: true,
        drawVerticalLine: false,
        horizontalInterval: yi,
        getDrawingHorizontalLine: (v) => const FlLine(color: border, strokeWidth: 1),
      );

  Widget barChart(List<MapEntry<String, double>> d) {
    if (d.isEmpty) return const Center(child: Text('لا توجد بيانات'));
    final r = yRange(d.map((e) => e.value).toList(), zero: true);
    final yi = (r[1] - r[0]) / 4;
    return BarChart(BarChartData(
      minY: r[0],
      maxY: r[1],
      gridData: grid(yi),
      borderData: FlBorderData(show: false),
      titlesData: titles(d, 1, yi),
      barTouchData: BarTouchData(
        touchTooltipData: BarTouchTooltipData(
          getTooltipColor: (g) => card2,
          getTooltipItem: (g, gi, rod, ri) => BarTooltipItem(
            '${d[g.x].key}\n${compact(rod.toY)}',
            const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 12),
          ),
        ),
      ),
      barGroups: [
        for (var i = 0; i < d.length; i++)
          BarChartGroupData(x: i, barRods: [
            BarChartRodData(
              toY: d[i].value,
              width: 18,
              borderRadius: const BorderRadius.vertical(top: Radius.circular(6)),
              gradient: LinearGradient(
                begin: Alignment.bottomCenter,
                end: Alignment.topCenter,
                colors: [blue.withOpacity(0.6), const Color(0xFF5B8CFF)],
              ),
            )
          ])
      ],
    ));
  }

  Widget lineChart(List<MapEntry<String, double>> d) {
    final interval = d.length > 8 ? (d.length / 6).ceilToDouble() : 1.0;
    final r = yRange(d.map((e) => e.value).toList(), zero: false);
    final yi = (r[1] - r[0]) / 4;
    return LineChart(LineChartData(
      minY: r[0],
      maxY: r[1],
      gridData: grid(yi),
      borderData: FlBorderData(show: false),
      titlesData: titles(d, interval, yi),
      lineTouchData: LineTouchData(
        touchTooltipData: LineTouchTooltipData(
          getTooltipColor: (s) => card2,
          getTooltipItems: (spots) => spots
              .map((p) => LineTooltipItem(
                    '${d[p.x.toInt()].key}\n${compact(p.y)}',
                    const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 12),
                  ))
              .toList(),
        ),
      ),
      lineBarsData: [
        LineChartBarData(
          spots: [for (var i = 0; i < d.length; i++) FlSpot(i.toDouble(), d[i].value)],
          isCurved: true,
          color: blue,
          barWidth: 3,
          dotData: const FlDotData(show: true),
          belowBarData: BarAreaData(
            show: true,
            gradient: LinearGradient(
              begin: Alignment.topCenter,
              end: Alignment.bottomCenter,
              colors: [blue.withOpacity(0.35), blue.withOpacity(0)],
            ),
          ),
        ),
      ],
    ));
  }

  Widget donut(List<MapEntry<String, double>> top) {
    final pos = top.where((e) => e.value > 0).toList();
    if (pos.isEmpty) return const Text('لا توجد بيانات موجبة', style: TextStyle(color: muted));
    final head = pos.take(5).toList();
    final rest = pos.skip(5).fold<double>(0, (a, e) => a + e.value);
    final items = [...head, if (rest > 0) MapEntry('أخرى', rest)];
    final total = items.fold<double>(0, (a, e) => a + e.value);
    return Row(children: [
      SizedBox(
        width: 150,
        height: 150,
        child: Stack(alignment: Alignment.center, children: [
          PieChart(PieChartData(
            sectionsSpace: 3,
            centerSpaceRadius: 46,
            sections: [
              for (var i = 0; i < items.length; i++)
                PieChartSectionData(
                  value: items[i].value,
                  color: palette[i % palette.length],
                  radius: 22,
                  showTitle: false,
                )
            ],
          )),
          Column(mainAxisSize: MainAxisSize.min, children: [
            Text(compact(total), style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
            const Text('الإجمالي', style: TextStyle(color: muted, fontSize: 10)),
          ]),
        ]),
      ),
      const SizedBox(width: 14),
      Expanded(
        child: Column(children: [
          for (var i = 0; i < items.length; i++)
            Padding(
              padding: const EdgeInsets.symmetric(vertical: 5),
              child: Row(children: [
                Container(width: 10, height: 10, decoration: BoxDecoration(color: palette[i % palette.length], shape: BoxShape.circle)),
                const SizedBox(width: 8),
                Expanded(child: Text(items[i].key, overflow: TextOverflow.ellipsis, style: const TextStyle(fontSize: 13))),
                Text('${(items[i].value / total * 100).toStringAsFixed(0)}%', style: const TextStyle(color: muted, fontSize: 12)),
              ]),
            ),
        ]),
      ),
    ]);
  }

  Widget ranked(List<MapEntry<String, double>> d) {
    if (d.isEmpty) return const Text('لا توجد بيانات', style: TextStyle(color: muted));
    final mx = d.map((e) => e.value.abs()).reduce((a, b) => a > b ? a : b);
    return Column(children: [
      for (var i = 0; i < d.length; i++)
        Padding(
          padding: const EdgeInsets.symmetric(vertical: 7),
          child: Row(children: [
            SizedBox(width: 90, child: Text(d[i].key, overflow: TextOverflow.ellipsis, style: const TextStyle(fontSize: 13))),
            Expanded(
              child: ClipRRect(
                borderRadius: BorderRadius.circular(8),
                child: LinearProgressIndicator(
                  value: mx == 0 ? 0 : (d[i].value.abs() / mx).clamp(0.0, 1.0),
                  minHeight: 10,
                  backgroundColor: card2,
                  valueColor: AlwaysStoppedAnimation(palette[i % palette.length]),
                ),
              ),
            ),
            const SizedBox(width: 10),
            SizedBox(width: 56, child: Text(compact(d[i].value), textAlign: TextAlign.end, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 12))),
          ]),
        ),
    ]);
  }

  List<List<String>> tableRows(Tbl s, int cols) {
    final key = '${identityHashCode(s)}|$q|$sortCol|$asc';
    if (key == _tKey) return _tRows;
    _tKey = key;
    var rows = <List<String>>[];
    if (q.isEmpty) {
      rows = List.of(s.rows);
    } else {
      final ql = q.toLowerCase();
      final hay = s.hay;
      for (var i = 0; i < hay.length; i++) {
        if (hay[i].contains(ql)) rows.add(s.rows[i]);
      }
    }
    if (sortCol != null && sortCol! < cols) {
      final c = sortCol!;
      if (s.numCols.contains(c)) {
        final keyed = rows.map((r) => MapEntry(toNum(r[c]) ?? -1e308, r)).toList();
        keyed.sort((x, y) => asc ? x.key.compareTo(y.key) : y.key.compareTo(x.key));
        rows = keyed.map((e) => e.value).toList();
      } else {
        rows.sort((x, y) => asc ? x[c].compareTo(y[c]) : y[c].compareTo(x[c]));
      }
    }
    return _tRows = rows;
  }

  Widget table(Tbl s) {
    final cols = math.min(s.head.length, 15);
    final rows = tableRows(s, cols);
    final pages = math.max(1, (rows.length / pageSize).ceil());
    if (page >= pages) page = pages - 1;
    final from = rows.isEmpty ? 0 : page * pageSize + 1;
    final to = math.min(rows.length, (page + 1) * pageSize);
    final slice = rows.skip(page * pageSize).take(pageSize).toList();

    return Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
      TextField(
        controller: qc,
        onChanged: (v) {
          _deb?.cancel();
          _deb = Timer(const Duration(milliseconds: 300), () {
            if (!mounted) return;
            setState(() {
              q = v.trim();
              page = 0;
            });
          });
        },
        decoration: InputDecoration(
          hintText: 'بحث في البيانات...',
          prefixIcon: const Icon(Icons.search),
          filled: true,
          fillColor: card2,
          isDense: true,
          border: OutlineInputBorder(borderRadius: BorderRadius.circular(12), borderSide: BorderSide.none),
        ),
      ),
      const SizedBox(height: 8),
      Text('عرض $from–$to من ${rows.length} صف${s.head.length > cols ? ' • أول $cols عمود' : ''}',
          style: const TextStyle(color: muted, fontSize: 12)),
      const SizedBox(height: 8),
      SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: DataTable(
          sortColumnIndex: (sortCol != null && sortCol! < cols) ? sortCol : null,
          sortAscending: asc,
          headingRowColor: WidgetStateProperty.all(card2),
          columnSpacing: 22,
          columns: [
            for (var i = 0; i < cols; i++)
              DataColumn(
                label: Text(s.head[i], style: const TextStyle(fontWeight: FontWeight.bold)),
                numeric: s.numCols.contains(i),
                onSort: (idx, a) => setState(() {
                  sortCol = idx;
                  asc = a;
                  page = 0;
                }),
              )
          ],
          rows: slice
              .map((r) => DataRow(cells: [
                    for (var i = 0; i < cols; i++)
                      DataCell(ConstrainedBox(
                        constraints: const BoxConstraints(maxWidth: 180),
                        child: Text(r[i], maxLines: 1, overflow: TextOverflow.ellipsis),
                      ))
                  ]))
              .toList(),
        ),
      ),
      if (pages > 1)
        Row(mainAxisAlignment: MainAxisAlignment.center, children: [
          IconButton(
            onPressed: page > 0 ? () => setState(() => page--) : null,
            icon: const Icon(Icons.arrow_back),
          ),
          Text('${page + 1} / $pages', style: const TextStyle(color: muted)),
          IconButton(
            onPressed: page < pages - 1 ? () => setState(() => page++) : null,
            icon: const Icon(Icons.arrow_forward),
          ),
        ]),
    ]);
  }
}

class Panel extends StatelessWidget {
  final String title;
  final IconData icon;
  final Widget child;
  final Widget? trailing;
  const Panel({super.key, required this.title, required this.icon, required this.child, this.trailing});

  @override
  Widget build(BuildContext context) => Container(
        margin: const EdgeInsets.only(bottom: 14),
        padding: const EdgeInsets.all(16),
        decoration: BoxDecoration(
          color: card,
          borderRadius: BorderRadius.circular(18),
          border: Border.all(color: border),
        ),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Row(children: [
            Container(
              padding: const EdgeInsets.all(7),
              decoration: BoxDecoration(color: blue.withOpacity(0.18), borderRadius: BorderRadius.circular(10)),
              child: Icon(icon, size: 18, color: const Color(0xFF5B8CFF)),
            ),
            const SizedBox(width: 10),
            Expanded(child: Text(title, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold))),
            if (trailing != null) trailing!,
          ]),
          const SizedBox(height: 14),
          child,
        ]),
      );
}

class Kpi extends StatelessWidget {
  final String title;
  final String value;
  final IconData icon;
  final List<Color> colors;
  const Kpi(this.title, this.value, this.icon, this.colors, {super.key});

  @override
  Widget build(BuildContext context) => Container(
        decoration: BoxDecoration(
          gradient: LinearGradient(colors: colors, begin: Alignment.topRight, end: Alignment.bottomLeft),
          borderRadius: BorderRadius.circular(18),
          boxShadow: [BoxShadow(color: colors.last.withOpacity(0.25), blurRadius: 14, offset: const Offset(0, 6))],
        ),
        child: Stack(children: [
          Positioned(left: -8, bottom: -8, child: Icon(icon, size: 78, color: Colors.white.withOpacity(0.13))),
          Positioned.fill(
            child: Padding(
              padding: const EdgeInsets.all(14),
              child: Column(crossAxisAlignment: CrossAxisAlignment.start, mainAxisAlignment: MainAxisAlignment.spaceBetween, children: [
                Container(
                  padding: const EdgeInsets.all(7),
                  decoration: BoxDecoration(color: Colors.white.withOpacity(0.22), shape: BoxShape.circle),
                  child: Icon(icon, size: 18, color: Colors.white),
                ),
                Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                  FittedBox(
                    fit: BoxFit.scaleDown,
                    child: Text(value, style: const TextStyle(fontSize: 24, fontWeight: FontWeight.bold, color: Colors.white)),
                  ),
                  Text(title, style: const TextStyle(color: Colors.white70, fontSize: 12)),
                ]),
              ]),
            ),
          ),
        ]),
      );
}