import 'dart:async';
import 'dart:convert';
import 'dart:math' as math;
import 'dart:typed_data';
import 'dart:ui' as ui;

import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:firebase_core/firebase_core.dart';
import 'package:google_sign_in/google_sign_in.dart';
import 'package:flutter/foundation.dart' show defaultTargetPlatform, kIsWeb;
import 'package:flutter/rendering.dart' show RenderRepaintBoundary;
import 'package:flutter/services.dart' show Clipboard, ClipboardData;
import 'package:file_picker/file_picker.dart';
import 'package:excel_plus/excel_plus.dart' as xls;
import 'package:image/image.dart' as img;
// ignore: deprecated_member_use
import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'package:intl/intl.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:printing/printing.dart';
import 'package:url_launcher/url_launcher.dart';

const _brand = Color(0xff2457a7);
const _ink = Color(0xff192b25);
const _paper = Color(0xfff4f7fc);
const _currencyNames = <String, String>{
  'NIO': 'Córdoba nicaragüense (C\$)',
  'USD': 'Dólar estadounidense (US\$)',
  'GTQ': 'Quetzal guatemalteco (Q)',
  'HNL': 'Lempira hondureño (L)',
  'CRC': 'Colón costarricense (₡)',
  'MXN': 'Peso mexicano (MX\$)',
  'EUR': 'Euro (€)',
};
const _currencySymbols = <String, String>{
  'NIO': 'C\$',
  'USD': '\$',
  'GTQ': 'Q',
  'HNL': 'L',
  'CRC': '₡',
  'MXN': 'MX\$',
  'EUR': '€'
};
const _currencyLocales = <String, String>{
  'NIO': 'es_NI',
  'USD': 'en_US',
  'GTQ': 'es_GT',
  'HNL': 'es_HN',
  'CRC': 'es_CR',
  'MXN': 'es_MX',
  'EUR': 'es_ES'
};
const _catalogColors = <String>[
  '#2457A7',
  '#2563EB',
  '#0891B2',
  '#0F766E',
  '#15803D',
  '#65A30D',
  '#A16207',
  '#EA580C',
  '#BE4E2A',
  '#DC2626',
  '#DB2777',
  '#C026D3',
  '#7C3AED',
  '#4F46E5',
  '#334155',
  '#111827',
  '#795548',
  '#D97706',
];
const _catalogThemes = <String, String>{
  'Moderno': 'modern',
  'Minimalista': 'minimal',
  'Editorial': 'editorial',
  'Creativo': 'playful',
  'Elegante': 'elegant',
};
bool _googleSignInInitialized = false;
const _googleClientId = String.fromEnvironment('GOOGLE_CLIENT_ID');
const _googleServerClientId = String.fromEnvironment('GOOGLE_SERVER_CLIENT_ID');
String _currentCurrency = 'NIO';
void _setCurrentCurrency(String value) {
  _currentCurrency = _currencyNames.containsKey(value) ? value : 'NIO';
}

String _money(double value) => NumberFormat.currency(
      locale: _currencyLocales[_currentCurrency],
      symbol: _currencySymbols[_currentCurrency],
      decimalDigits: 2,
    ).format(value);

DateTime _addMonths(DateTime date, int months) {
  final monthIndex = date.month - 1 + months;
  final year = date.year + monthIndex ~/ 12;
  final month = monthIndex % 12 + 1;
  final lastDay = DateTime(year, month + 1, 0).day;
  return DateTime(
      year, month, math.min(date.day, lastDay), date.hour, date.minute);
}

List<Map<String, dynamic>> _buildPaymentSchedule({
  required double balance,
  required int count,
  required String frequency,
  required DateTime firstDueDate,
}) {
  final cents = (balance * 100).round();
  if (cents <= 0) return [];
  final installments = math.min(count.clamp(1, 12), cents);
  final baseCents = cents ~/ installments;
  final extraCents = cents % installments;
  return List.generate(installments, (index) {
    final dueDate = switch (frequency) {
      'weekly' => firstDueDate.add(Duration(days: 7 * index)),
      'biweekly' => firstDueDate.add(Duration(days: 14 * index)),
      _ => _addMonths(firstDueDate, index),
    };
    final amount = (baseCents + (index < extraCents ? 1 : 0)) / 100;
    return {
      'number': index + 1,
      'amount': amount,
      'paidAmount': 0.0,
      'remainingAmount': amount,
      'dueAt': DateTime(dueDate.year, dueDate.month, dueDate.day, 9)
          .toUtc()
          .toIso8601String(),
      'status': 'pending',
    };
  });
}

List<Map<String, dynamic>> _applyInstallmentPayment(
    dynamic rawSchedule, double amount) {
  if (rawSchedule is! List) return [];
  var unapplied = (amount * 100).round();
  final schedule = rawSchedule
      .whereType<Map>()
      .map((item) => Map<String, dynamic>.from(item))
      .toList();
  for (final item in schedule) {
    if (unapplied <= 0) break;
    final remainingCents = (_number(item['remainingAmount'] ??
                (_number(item['amount']) - _number(item['paidAmount']))) *
            100)
        .round();
    if (remainingCents <= 0) {
      item['status'] = 'paid';
      item['remainingAmount'] = 0.0;
      continue;
    }
    final appliedCents = math.min(unapplied, remainingCents);
    final paidCents =
        (_number(item['paidAmount']) * 100).round() + appliedCents;
    final leftCents = remainingCents - appliedCents;
    item['paidAmount'] = paidCents / 100;
    item['remainingAmount'] = leftCents / 100;
    item['status'] = leftCents == 0 ? 'paid' : 'partial';
    unapplied -= appliedCents;
  }
  return schedule;
}

double _roundMoney(double value) => (value * 100).round() / 100;

Map<String, dynamic> _calculateSaleReturn({
  required Map<String, dynamic> sale,
  required List<Map<String, dynamic>> requestedItems,
  required Map<String, Map<String, dynamic>> productsById,
  required String reason,
  required String refundMethod,
  required String returnId,
  required DateTime at,
}) {
  final soldQuantities = <String, double>{};
  final soldValues = <String, double>{};
  final soldNames = <String, String>{};
  final saleItems = sale['items'] is List
      ? (sale['items'] as List).whereType<Map>()
      : const <Map>[];
  for (final item in saleItems) {
    final id = item['productId']?.toString() ?? '';
    final quantity = _number(item['quantity']);
    if (id.isEmpty || quantity <= 0) continue;
    final value = _number(item['total']);
    final unitPrice = item['unitPrice'] == null
        ? (quantity > 0 ? value / quantity : 0)
        : _number(item['unitPrice']);
    soldQuantities.update(id, (current) => current + quantity,
        ifAbsent: () => quantity);
    soldValues.update(id, (current) => current + unitPrice * quantity,
        ifAbsent: () => unitPrice * quantity);
    soldNames.putIfAbsent(id, () => item['name']?.toString() ?? 'Producto');
  }
  if (soldQuantities.isEmpty)
    throw Exception('La venta no tiene productos que devolver.');

  final returnedRaw = sale['returnedQuantities'];
  final returnedQuantities = returnedRaw is Map
      ? returnedRaw
          .map((key, value) => MapEntry(key.toString(), _number(value)))
      : <String, double>{};
  final requestedQuantities = <String, double>{};
  for (final item in requestedItems) {
    final id = item['productId']?.toString() ?? '';
    final quantity = _number(item['quantity']);
    if (id.isEmpty || quantity <= 0) continue;
    requestedQuantities.update(id, (current) => current + quantity,
        ifAbsent: () => quantity);
  }
  if (requestedQuantities.isEmpty)
    throw Exception('Selecciona al menos una cantidad para devolver.');

  final returnLines = <Map<String, dynamic>>[];
  var returnSubtotal = 0.0;
  for (final entry in requestedQuantities.entries) {
    final sold = soldQuantities[entry.key] ?? 0;
    final alreadyReturned = returnedQuantities[entry.key] ?? 0;
    if (!productsById.containsKey(entry.key)) {
      throw Exception(
          'No encontramos el producto ${soldNames[entry.key] ?? ''} en el inventario.');
    }
    if (entry.value > sold - alreadyReturned + .000001) {
      throw Exception(
          'La cantidad a devolver supera las unidades vendidas que quedan disponibles.');
    }
    final unitPrice = sold <= 0 ? 0 : soldValues[entry.key]! / sold;
    final lineTotal = _roundMoney(unitPrice * entry.value);
    returnSubtotal += lineTotal;
    returnedQuantities[entry.key] = alreadyReturned + entry.value;
    returnLines.add({
      'productId': entry.key,
      'name': soldNames[entry.key] ??
          productsById[entry.key]?['name'] ??
          'Producto',
      'quantity': entry.value,
      'unitPrice': unitPrice,
      'total': lineTotal,
    });
  }
  returnSubtotal = _roundMoney(returnSubtotal);

  final originalSubtotal = _number(sale['subtotal']) > 0
      ? _number(sale['subtotal'])
      : saleItems.fold<double>(0, (sum, item) => sum + _number(item['total']));
  final originalPocketFee = _number(sale['pocketSurcharge']);
  final oldReturnedSubtotal = _number(sale['returnedItemsSubtotal']);
  final oldReturnedFee = _number(sale['returnedPocketSurcharge']);
  final newReturnedSubtotal = _roundMoney(oldReturnedSubtotal + returnSubtotal);
  final newReturnedFee = originalPocketFee > 0 && originalSubtotal > 0
      ? _roundMoney(
          originalPocketFee * (newReturnedSubtotal / originalSubtotal))
      : 0.0;
  final pocketFeeRefund = _roundMoney(newReturnedFee - oldReturnedFee);
  final returnTotal = _roundMoney(returnSubtotal + pocketFeeRefund);
  final currentBalance = _number(sale['balance']);
  final balanceAdjusted = _roundMoney(math.min(currentBalance, returnTotal));
  final refundAmount = _roundMoney(returnTotal - balanceAdjusted);
  final currentPaid = _number(sale['amountPaid']);
  if (refundAmount > currentPaid + .01) {
    throw Exception(
        'El reembolso supera el importe pagado de la venta. Actualiza los datos antes de continuar.');
  }
  if (refundAmount > 0 && refundMethod.trim().isEmpty) {
    throw Exception('Selecciona cómo se devolverá el importe al cliente.');
  }
  final currentTotal = _number(sale['total']);
  if (returnTotal <= 0 || returnTotal > currentTotal + .01) {
    throw Exception('El importe de la devolución no es válido.');
  }
  final nextBalance =
      _roundMoney(math.max(0, currentBalance - balanceAdjusted));
  final nextPaid = _roundMoney(math.max(0, currentPaid - refundAmount));
  final date = DateFormat('yyyy-MM-dd HH:mm').format(at);
  final adjustedSchedule =
      _applyInstallmentPayment(sale['paymentSchedule'], balanceAdjusted);
  final saleUpdate = <String, dynamic>{
    'total': _roundMoney(currentTotal - returnTotal),
    'amountPaid': nextPaid,
    'balance': nextBalance,
    'paymentStatus': nextBalance <= .005 ? 'Pagada' : 'Pendiente',
    'autoReminders': nextBalance > .005 && sale['autoReminders'] == true,
    'returnedQuantities': returnedQuantities,
    'returnedItemsSubtotal': newReturnedSubtotal,
    'returnedPocketSurcharge': newReturnedFee,
    'returnedAmount':
        _roundMoney(_number(sale['returnedAmount']) + returnTotal),
    'returnsCount': _number(sale['returnsCount']).toInt() + 1,
    'lastReturnAt': date,
    'paymentSchedule': adjustedSchedule,
  };
  final returnRecord = <String, dynamic>{
    'id': returnId,
    'saleId': sale['id']?.toString() ?? '',
    'saleNumber': sale['number']?.toString() ?? '',
    'customer': sale['party']?.toString() ?? '',
    'items': returnLines,
    'subtotal': returnSubtotal,
    'pocketSurchargeRefund': pocketFeeRefund,
    'total': returnTotal,
    'balanceAdjusted': balanceAdjusted,
    'refundAmount': refundAmount,
    'refundMethod': refundAmount > 0 ? refundMethod : 'Ajuste de saldo',
    'reason': reason.trim(),
    'currency': sale['currency']?.toString() ?? 'NIO',
    'date': date,
  };
  return {
    'saleUpdate': saleUpdate,
    'returnRecord': returnRecord,
    'stockAdditions': requestedQuantities,
  };
}

const _sections = <String>[
  'Resumen',
  'Ventas',
  'Compras',
  'Inventario',
  'Clientes',
  'Gastos',
  'WhatsApp',
  'Poket',
  'Catálogo web',
  'Configuración'
];
const _icons = <IconData>[
  Icons.space_dashboard_outlined,
  Icons.point_of_sale_outlined,
  Icons.shopping_bag_outlined,
  Icons.inventory_2_outlined,
  Icons.people_outline,
  Icons.receipt_long_outlined,
  Icons.chat_outlined,
  Icons.account_balance_wallet_outlined,
  Icons.storefront_outlined,
  Icons.settings_outlined,
];

String _resolveApiRoot(String configured) {
  final value = configured.trim();
  if (value.isNotEmpty) return value.replaceFirst(RegExp(r'/+$'), '');
  if (!kIsWeb) return '';
  final uri = Uri.base;
  if (['localhost', '127.0.0.1'].contains(uri.host) && uri.port != 3030) {
    return 'http://127.0.0.1:3030';
  }
  return uri.origin;
}

String _catalogSlugFromName(String value) {
  var slug = value.trim().toLowerCase();
  const accented = 'áàäâéèëêíìïîóòöôúùüûñç';
  const plain = 'aaaaeeeeiiiioooouuuunc';
  for (var index = 0; index < accented.length; index++) {
    slug = slug.replaceAll(accented[index], plain[index]);
  }
  final normalized = slug
      .replaceAll(RegExp(r'[^a-z0-9]+'), '-')
      .replaceAll(RegExp(r'^-+|-+$'), '')
      .replaceAll(RegExp(r'-+'), '-');
  return normalized.length <= 48
      ? normalized
      : normalized.substring(0, 48).replaceFirst(RegExp(r'-+$'), '');
}

String? _catalogSlugFromUrl(String? value) {
  final uri = Uri.tryParse(value ?? '');
  final segments = uri?.pathSegments ?? const <String>[];
  if (segments.length >= 3 && segments[segments.length - 3] == 'catalog') {
    return segments.last;
  }
  if (segments.length >= 2 && segments[segments.length - 2] == 'catalog') {
    return segments.last;
  }
  return null;
}

Future<Map<String, String>> _authorizedApiHeaders({bool json = false}) async {
  final user = FirebaseAuth.instance.currentUser;
  final token = await user?.getIdToken();
  if (token == null || token.isEmpty) {
    throw StateError('Inicia sesión para acceder a los datos de tu cuenta.');
  }
  return {
    'Authorization': 'Bearer $token',
    if (json) 'Content-Type': 'application/json',
  };
}

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  final options = _firebaseOptionsFromDefines();
  final firebaseEnabled = options != null &&
      (kIsWeb || defaultTargetPlatform != TargetPlatform.linux);
  if (firebaseEnabled) {
    await Firebase.initializeApp(options: options);
  }
  runApp(BusinessApp(firebaseEnabled: firebaseEnabled));
}

FirebaseOptions? _firebaseOptionsFromDefines() {
  const apiKey = String.fromEnvironment('FIREBASE_API_KEY',
      defaultValue: 'AIzaSyAfghXHozzPTgKn1MPSLIcgpk78GfCWE4c');
  const appId = String.fromEnvironment('FIREBASE_APP_ID',
      defaultValue: '1:173074046860:web:074d9ce4217a5f07bdc0a9');
  const sender = String.fromEnvironment('FIREBASE_MESSAGING_SENDER_ID',
      defaultValue: '173074046860');
  const project =
      String.fromEnvironment('FIREBASE_PROJECT_ID', defaultValue: 'cmr-boot');
  const authDomain = String.fromEnvironment('FIREBASE_AUTH_DOMAIN',
      defaultValue: 'cmr-boot.firebaseapp.com');
  const storageBucket = String.fromEnvironment('FIREBASE_STORAGE_BUCKET',
      defaultValue: 'cmr-boot.firebasestorage.app');
  const measurementId = String.fromEnvironment('FIREBASE_MEASUREMENT_ID',
      defaultValue: 'G-FJ19B2E00M');
  if ([apiKey, appId, sender, project].any((value) => value.isEmpty))
    return null;
  return FirebaseOptions(
    apiKey: apiKey,
    appId: appId,
    messagingSenderId: sender,
    projectId: project,
    authDomain: authDomain.isEmpty ? null : authDomain,
    storageBucket: storageBucket.isEmpty ? null : storageBucket,
    measurementId: measurementId.isEmpty ? null : measurementId,
  );
}

class BusinessApp extends StatelessWidget {
  const BusinessApp({super.key, required this.firebaseEnabled});
  final bool firebaseEnabled;
  @override
  Widget build(BuildContext context) => MaterialApp(
        title: 'Sync Connect',
        debugShowCheckedModeBanner: false,
        theme: ThemeData(
          useMaterial3: true,
          scaffoldBackgroundColor: _paper,
          colorScheme:
              ColorScheme.fromSeed(seedColor: _brand, surface: Colors.white),
          fontFamily: 'Arial',
          inputDecorationTheme: InputDecorationTheme(
            filled: true,
            fillColor: Colors.white,
            border: OutlineInputBorder(
                borderRadius: BorderRadius.circular(12),
                borderSide: const BorderSide(color: Color(0xffe2e8e3))),
            enabledBorder: OutlineInputBorder(
                borderRadius: BorderRadius.circular(12),
                borderSide: const BorderSide(color: Color(0xffe2e8e3))),
          ),
        ),
        home: firebaseEnabled ? const AuthGate() : const AuthUnavailablePage(),
      );
}

class AuthGate extends StatefulWidget {
  const AuthGate({super.key});
  @override
  State<AuthGate> createState() => _AuthGateState();
}

class _AuthGateState extends State<AuthGate> {
  late Stream<User?> _authStream;

  // Firebase may take longer than a few seconds to restore a persisted
  // session, especially on a cold start. A Stream.timeout here turned that
  // delay into a false "session validation" error even when the device was
  // online. Let Firebase emit its initial auth state when it is ready.
  Stream<User?> _newAuthStream() => FirebaseAuth.instance.authStateChanges();

  @override
  void initState() {
    super.initState();
    _authStream = _newAuthStream();
  }

  @override
  Widget build(BuildContext context) => StreamBuilder<User?>(
        stream: _authStream,
        builder: (context, snapshot) {
          if (snapshot.connectionState == ConnectionState.waiting)
            return const _Loading(message: 'Comprobando tu sesión…');
          if (snapshot.hasError) {
            return _AuthRetryPage(
                onRetry: () => setState(() => _authStream = _newAuthStream()));
          }
          final user = snapshot.data ?? FirebaseAuth.instance.currentUser;
          if (user == null) return const SignInPage();
          return BusinessShell(
              key: ValueKey(user.uid),
              userId: user.uid,
              email: user.email ?? '',
              demo: false);
        },
      );
}

class _AuthRetryPage extends StatelessWidget {
  const _AuthRetryPage({required this.onRetry});
  final VoidCallback onRetry;
  @override
  Widget build(BuildContext context) => Scaffold(
        body: Center(
          child: Padding(
            padding: const EdgeInsets.all(24),
            child: Column(mainAxisSize: MainAxisSize.min, children: [
              const Icon(Icons.cloud_sync_outlined, size: 48, color: _brand),
              const SizedBox(height: 12),
              const Text('No pudimos restaurar tu sesión',
                  style: TextStyle(fontSize: 20, fontWeight: FontWeight.w700)),
              const SizedBox(height: 8),
              const Text(
                  'Vuelve a intentar o inicia sesión de nuevo. Tu conexión a Internet puede estar bien.',
                  textAlign: TextAlign.center),
              const SizedBox(height: 16),
              FilledButton.icon(
                  onPressed: onRetry,
                  icon: const Icon(Icons.refresh),
                  label: const Text('Reintentar')),
            ]),
          ),
        ),
      );
}

class AuthUnavailablePage extends StatelessWidget {
  const AuthUnavailablePage({super.key});
  @override
  Widget build(BuildContext context) => Scaffold(
        body: Center(
          child: Padding(
            padding: EdgeInsets.all(28),
            child: ConstrainedBox(
              constraints: BoxConstraints(maxWidth: 460),
              child: Column(mainAxisSize: MainAxisSize.min, children: [
                _BrandMark(),
                const SizedBox(height: 20),
                const Text('Acceso con Google requerido',
                    style: TextStyle(fontSize: 24, fontWeight: FontWeight.w800),
                    textAlign: TextAlign.center),
                const SizedBox(height: 10),
                const Text(
                    'Esta versión no tiene Firebase habilitado para iniciar sesión. Configura Firebase para continuar.',
                    textAlign: TextAlign.center,
                    style: TextStyle(color: Colors.black54)),
              ]),
            ),
          ),
        ),
      );
}

class SignInPage extends StatefulWidget {
  const SignInPage({super.key});
  @override
  State<SignInPage> createState() => _SignInPageState();
}

class _SignInPageState extends State<SignInPage> {
  final _email = TextEditingController();
  final _password = TextEditingController();
  bool _busy = false;
  bool _register = false;
  bool _hidePassword = true;
  String? _error;
  String? _message;
  @override
  void dispose() {
    _email.dispose();
    _password.dispose();
    super.dispose();
  }

  Future<void> _submitEmail() async {
    setState(() {
      _busy = true;
      _error = null;
      _message = null;
    });
    try {
      if (_register) {
        await FirebaseAuth.instance.createUserWithEmailAndPassword(
            email: _email.text.trim(), password: _password.text);
      } else {
        await FirebaseAuth.instance.signInWithEmailAndPassword(
            email: _email.text.trim(), password: _password.text);
      }
    } on FirebaseAuthException catch (e) {
      setState(() => _error = switch (e.code) {
            'email-already-in-use' =>
              'Ese correo ya tiene una cuenta. Inicia sesión.',
            'weak-password' =>
              'La contraseña debe tener al menos 6 caracteres.',
            'invalid-email' => 'Revisa el formato del correo electrónico.',
            'invalid-credential' ||
            'user-not-found' ||
            'wrong-password' =>
              'Correo o contraseña incorrectos.',
            'operation-not-allowed' =>
              'Activa Correo electrónico/contraseña en Firebase Authentication.',
            _ => e.message ?? 'No se pudo completar el acceso.',
          });
    } catch (e) {
      setState(() => _error = 'No se pudo conectar con Firebase: $e');
    }
    if (mounted) setState(() => _busy = false);
  }

  Future<void> _sendPasswordReset() async {
    final email = _email.text.trim();
    if (email.isEmpty || !email.contains('@')) {
      setState(() {
        _error = 'Escribe primero el correo de tu cuenta.';
        _message = null;
      });
      return;
    }
    setState(() {
      _busy = true;
      _error = null;
      _message = null;
    });
    try {
      await FirebaseAuth.instance.sendPasswordResetEmail(email: email);
      if (mounted) {
        setState(() => _message =
            'Si existe una cuenta con ese correo, Firebase enviará un enlace para restablecer la contraseña. Revisa también Spam.');
      }
    } on FirebaseAuthException catch (error) {
      if (!mounted) return;
      setState(() => _error = switch (error.code) {
            'invalid-email' => 'Revisa el formato del correo electrónico.',
            'operation-not-allowed' =>
              'Activa Correo electrónico/contraseña en Firebase Authentication.',
            'too-many-requests' =>
              'Se hicieron varios intentos. Espera un momento y vuelve a intentarlo.',
            'network-request-failed' =>
              'No se pudo conectar con Firebase. Intenta de nuevo.',
            _ =>
              'No se pudo enviar el enlace. Revisa el correo e inténtalo de nuevo.',
          });
    } catch (_) {
      if (mounted)
        setState(() => _error =
            'No se pudo enviar el enlace. Inténtalo de nuevo más tarde.');
    } finally {
      if (mounted) setState(() => _busy = false);
    }
  }

  Future<void> _signInWithGoogle() async {
    setState(() {
      _busy = true;
      _error = null;
      _message = null;
    });
    try {
      final desktop = !kIsWeb &&
          [TargetPlatform.windows, TargetPlatform.linux]
              .contains(defaultTargetPlatform);
      if (desktop) {
        const configuredWebUrl = String.fromEnvironment('PUBLIC_WEB_URL');
        const configuredApiUrl = String.fromEnvironment('BOT_API_URL');
        final webUri = Uri.tryParse((configuredWebUrl.trim().isNotEmpty
                ? configuredWebUrl
                : configuredApiUrl)
            .trim());
        if (webUri == null ||
            !['https', 'http'].contains(webUri.scheme) ||
            webUri.host.isEmpty) {
          setState(() => _error =
              'Configura PUBLIC_WEB_URL con la dirección HTTPS de tu versión web para iniciar sesión con Google.');
        } else {
          await Clipboard.setData(ClipboardData(text: webUri.toString()));
          final opened = await launchUrl(
            webUri,
            mode: LaunchMode.externalApplication,
          );
          if (opened) {
            setState(() => _message =
                'Abrimos la versión web en tu navegador y copiamos el enlace. Inicia sesión con Google allí; esta aplicación de escritorio requiere iniciar sesión con correo y contraseña.');
          } else {
            setState(() => _error =
                'No se pudo abrir el navegador. El enlace de la versión web quedó copiado.');
          }
        }
      } else if (kIsWeb) {
        await FirebaseAuth.instance.signInWithPopup(GoogleAuthProvider());
      } else if ([
        TargetPlatform.android,
        TargetPlatform.iOS,
        TargetPlatform.macOS
      ].contains(defaultTargetPlatform)) {
        if (!_googleSignInInitialized) {
          await GoogleSignIn.instance.initialize(
            clientId: _googleClientId.isEmpty ? null : _googleClientId,
            serverClientId:
                _googleServerClientId.isEmpty ? null : _googleServerClientId,
          );
          _googleSignInInitialized = true;
        }
        final account = await GoogleSignIn.instance.authenticate();
        final credential = GoogleAuthProvider.credential(
          idToken: account.authentication.idToken,
        );
        await FirebaseAuth.instance.signInWithCredential(credential);
      }
    } on FirebaseAuthException catch (e) {
      setState(() => _error = e.code == 'operation-not-allowed'
          ? 'Activa Google como proveedor de acceso en Firebase Authentication.'
          : e.message ?? 'No se pudo iniciar sesión con Google.');
    } on GoogleSignInException catch (e) {
      if (e.code != GoogleSignInExceptionCode.canceled) {
        setState(() =>
            _error = e.description ?? 'No se pudo iniciar sesión con Google.');
      }
    } catch (e) {
      setState(() => _error = 'No se pudo completar el acceso con Google: $e');
    }
    if (mounted) setState(() => _busy = false);
  }

  @override
  Widget build(BuildContext context) => Scaffold(
        body: Center(
            child: ConstrainedBox(
                constraints: const BoxConstraints(maxWidth: 460),
                child: Card(
                  margin: const EdgeInsets.all(24),
                  elevation: 0,
                  shape: RoundedRectangleBorder(
                      borderRadius: BorderRadius.circular(24),
                      side: const BorderSide(color: Color(0xffe5ebe6))),
                  child: Padding(
                      padding: const EdgeInsets.all(32),
                      child: Column(
                          mainAxisSize: MainAxisSize.min,
                          crossAxisAlignment: CrossAxisAlignment.stretch,
                          children: [
                            const _BrandMark(),
                            const SizedBox(height: 24),
                            Text(_register ? 'Crea tu cuenta' : 'Bienvenido',
                                style: const TextStyle(
                                    fontSize: 28,
                                    fontWeight: FontWeight.w800,
                                    color: _ink)),
                            const SizedBox(height: 8),
                            Text(
                                _register
                                    ? 'Regístrate con Google o con tu correo electrónico.'
                                    : 'Inicia sesión con Google o con tu correo electrónico.',
                                style: const TextStyle(color: Colors.black54)),
                            const SizedBox(height: 24),
                            TextField(
                              controller: _email,
                              keyboardType: TextInputType.emailAddress,
                              autofillHints: const [AutofillHints.email],
                              decoration: const InputDecoration(
                                  labelText: 'Correo electrónico',
                                  prefixIcon: Icon(Icons.mail_outline)),
                            ),
                            const SizedBox(height: 12),
                            TextField(
                              controller: _password,
                              obscureText: _hidePassword,
                              autofillHints: [
                                _register
                                    ? AutofillHints.newPassword
                                    : AutofillHints.password
                              ],
                              onSubmitted: (_) => _submitEmail(),
                              decoration: InputDecoration(
                                labelText: 'Contraseña',
                                prefixIcon: const Icon(Icons.lock_outline),
                                suffixIcon: IconButton(
                                  tooltip: _hidePassword
                                      ? 'Mostrar contraseña'
                                      : 'Ocultar contraseña',
                                  onPressed: () => setState(
                                      () => _hidePassword = !_hidePassword),
                                  icon: Icon(_hidePassword
                                      ? Icons.visibility_outlined
                                      : Icons.visibility_off_outlined),
                                ),
                              ),
                            ),
                            if (!_register)
                              Align(
                                alignment: Alignment.centerRight,
                                child: TextButton(
                                  onPressed: _busy ? null : _sendPasswordReset,
                                  child:
                                      const Text('¿Olvidaste tu contraseña?'),
                                ),
                              ),
                            if (_error != null)
                              Padding(
                                  padding: const EdgeInsets.only(top: 12),
                                  child: Text(_error!,
                                      style:
                                          const TextStyle(color: Colors.red))),
                            if (_message != null)
                              Padding(
                                padding: const EdgeInsets.only(top: 12),
                                child: Text(_message!,
                                    style: const TextStyle(color: _brand)),
                              ),
                            const SizedBox(height: 18),
                            FilledButton(
                              onPressed: _busy ? null : _submitEmail,
                              child: Padding(
                                padding:
                                    const EdgeInsets.symmetric(vertical: 12),
                                child: Text(_busy
                                    ? 'Un momento…'
                                    : _register
                                        ? 'Crear cuenta con correo'
                                        : 'Iniciar sesión con correo'),
                              ),
                            ),
                            const Padding(
                              padding: EdgeInsets.symmetric(vertical: 8),
                              child: Row(children: [
                                Expanded(child: Divider()),
                                Padding(
                                    padding:
                                        EdgeInsets.symmetric(horizontal: 12),
                                    child: Text('o')),
                                Expanded(child: Divider())
                              ]),
                            ),
                            FilledButton.icon(
                              onPressed: _busy ? null : _signInWithGoogle,
                              icon: _busy
                                  ? const SizedBox(
                                      width: 20,
                                      height: 20,
                                      child: CircularProgressIndicator(
                                          strokeWidth: 2))
                                  : const Text('G',
                                      style: TextStyle(
                                          fontSize: 20,
                                          fontWeight: FontWeight.w800)),
                              label: Padding(
                                padding:
                                    const EdgeInsets.symmetric(vertical: 12),
                                child: Text(_busy
                                    ? 'Conectando con Google…'
                                    : (!kIsWeb &&
                                            [
                                              TargetPlatform.windows,
                                              TargetPlatform.linux
                                            ].contains(defaultTargetPlatform)
                                        ? 'Abrir Google en navegador'
                                        : 'Continuar con Google')),
                              ),
                            ),
                            TextButton(
                              onPressed: _busy
                                  ? null
                                  : () => setState(() {
                                        _register = !_register;
                                        _error = null;
                                      }),
                              child: Text(_register
                                  ? 'Ya tengo cuenta · Iniciar sesión'
                                  : 'Crear una cuenta nueva'),
                            ),
                          ])),
                ))),
      );
}

class BusinessShell extends StatefulWidget {
  const BusinessShell(
      {super.key,
      required this.userId,
      required this.email,
      required this.demo});
  final String userId, email;
  final bool demo;
  @override
  State<BusinessShell> createState() => _BusinessShellState();
}

class _BusinessShellState extends State<BusinessShell> {
  int _selected = 0;
  late final _Store _store = _Store(widget.userId, widget.demo);
  @override
  void initState() {
    super.initState();
    unawaited(_store.load().then((_) => _reconcilePendingPocketPayments()));
  }

  @override
  Widget build(BuildContext context) => AnimatedBuilder(
        animation: _store,
        builder: (context, _) {
          return Scaffold(
            drawer: _drawer(),
            appBar: AppBar(
              backgroundColor: Colors.white,
              surfaceTintColor: Colors.white,
              leading: Builder(
                  builder: (context) => IconButton(
                        tooltip: 'Abrir menú',
                        icon: const Icon(Icons.menu_rounded),
                        onPressed: () => Scaffold.of(context).openDrawer(),
                      )),
              title: const Text('Sync Connect',
                  style: TextStyle(fontWeight: FontWeight.w800)),
              actions: [
                if (!widget.demo)
                  IconButton(
                      tooltip: 'Cerrar sesión',
                      onPressed: () async {
                        await FirebaseAuth.instance.signOut();
                        _store.clearSessionView();
                      },
                      icon: const Icon(Icons.logout)),
                const SizedBox(width: 8),
              ],
            ),
            body: Column(children: [
              if (_store.refreshing)
                const LinearProgressIndicator(minHeight: 2),
              Expanded(
                child: AnimatedSwitcher(
                  duration: const Duration(milliseconds: 240),
                  switchInCurve: Curves.easeOutCubic,
                  switchOutCurve: Curves.easeInCubic,
                  transitionBuilder: (child, animation) => FadeTransition(
                    opacity: animation,
                    child: SlideTransition(
                      position: animation.drive(
                          Tween(begin: const Offset(.015, 0), end: Offset.zero)
                              .chain(CurveTween(curve: Curves.easeOutCubic))),
                      child: child,
                    ),
                  ),
                  child: KeyedSubtree(key: ValueKey(_selected), child: _page()),
                ),
              ),
            ]),
            floatingActionButton: ![1, 2, 3, 4, 5].contains(_selected)
                ? null
                : FloatingActionButton.extended(
                    onPressed: () => _createForSection(),
                    icon: const Icon(Icons.add),
                    label: Text(_selected == 1
                        ? 'Nueva venta'
                        : _selected == 2
                            ? 'Nueva compra'
                            : _selected == 3
                                ? 'Nuevo producto'
                                : _selected == 4
                                    ? 'Nuevo cliente'
                                    : 'Nuevo gasto'),
                  ),
          );
        },
      );
  Widget _drawer() => Drawer(
          child: SafeArea(
              child: Column(children: [
        const Padding(
            padding: EdgeInsets.fromLTRB(20, 18, 20, 22),
            child: Align(alignment: Alignment.centerLeft, child: _BrandMark())),
        Expanded(
            child: ListView(
                padding: const EdgeInsets.symmetric(horizontal: 10),
                children: [
              for (var i = 0; i < _sections.length; i++)
                Padding(
                    padding: const EdgeInsets.symmetric(vertical: 2),
                    child: ListTile(
                        selected: _selected == i,
                        selectedColor: _brand,
                        selectedTileColor: const Color(0xffedf3ff),
                        shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(12)),
                        leading: Icon(_icons[i]),
                        title: Text(_sections[i],
                            style:
                                const TextStyle(fontWeight: FontWeight.w600)),
                        onTap: () {
                          Navigator.pop(context);
                          setState(() => _selected = i);
                        }))
            ])),
        const Divider(height: 1),
        ListTile(
            leading: const CircleAvatar(
                backgroundColor: Color(0xffedf3ff),
                child: Icon(Icons.person_outline, color: _brand)),
            title: Text(widget.email,
                maxLines: 1, overflow: TextOverflow.ellipsis),
            subtitle: Text(
                widget.demo ? 'Datos en este navegador' : 'Cuenta personal'))
      ])));

  Widget _page() => switch (_selected) {
        0 => _Dashboard(
            store: _store, onNavigate: (i) => setState(() => _selected = i)),
        1 => _RecordsPage(
            title: 'Ventas',
            subtitle: 'Ventas registradas y su estado de pago.',
            icon: _icons[1],
            records: _store.sales,
            store: _store,
            money: true),
        2 => _RecordsPage(
            title: 'Compras',
            subtitle: 'Entradas de mercancía y compras a proveedores.',
            icon: _icons[2],
            records: _store.purchases,
            money: true),
        3 => _ProductsPage(store: _store),
        4 => _ContactsPage(store: _store),
        5 => _RecordsPage(
            title: 'Gastos',
            subtitle: 'Otros egresos del negocio.',
            icon: _icons[5],
            records: _store.expenses,
            money: true),
        6 => const _WhatsAppPage(),
        7 => const _PoketPage(),
        8 => _CatalogPage(store: _store),
        _ => _ConnectionsPage(
            store: _store,
            onNavigate: (index) => setState(() => _selected = index)),
      };
  Future<void> _createForSection() async {
    if (_selected == 3) return _productDialog();
    if (_selected == 4) return _contactDialog();
    if (_selected == 1 || _selected == 2)
      return _transactionDialog(purchase: _selected == 2);
    return _expenseDialog();
  }

  Future<void> _productDialog([Map<String, dynamic>? existing]) async {
    final result = await showDialog<Map<String, dynamic>>(
        context: context, builder: (_) => _ProductDialog(initial: existing));
    if (result == null) return;
    try {
      await _store.saveProduct(result, id: existing?['id'] as String?);
      if (mounted) _notice('Producto guardado en Firebase.');
    } catch (error) {
      if (mounted) _notice('$error'.replaceFirst('Exception: ', ''));
    }
  }

  Future<void> _contactDialog([Map<String, dynamic>? existing]) async {
    final result = await showDialog<Map<String, dynamic>>(
        context: context, builder: (_) => _ContactDialog(initial: existing));
    if (result == null) return;
    try {
      await _store.saveContact(result, id: existing?['id'] as String?);
      if (mounted) _notice('Cliente guardado en Firebase.');
    } catch (error) {
      if (mounted) _notice('$error'.replaceFirst('Exception: ', ''));
    }
  }

  Future<void> _transactionDialog({required bool purchase}) async {
    if (_store.products.isEmpty) {
      _notice('Agrega un producto al inventario primero.');
      setState(() => _selected = 3);
      return;
    }
    final result = await showDialog<Map<String, dynamic>>(
        context: context,
        builder: (_) => _TransactionDialog(
            products: _store.products,
            contacts: _store.contacts,
            purchase: purchase));
    if (result == null) return;
    try {
      result['currency'] = _store.currencyCode;
      if (!purchase && result['paymentMethod'] == 'pocket') {
        await _createAndShowSalePaylink(result);
        return;
      }
      final remindersScheduled =
          await _store.recordTransaction(result, purchase: purchase);
      if (!purchase) {
        final sale = _store.sales.firstWhere(
          (item) => item['id']?.toString() == result['id']?.toString(),
          orElse: () => result,
        );
        await _showReceiptOptions(sale);
      } else {
        final purchaseRecord = _store.purchases.firstWhere(
          (item) => item['id']?.toString() == result['id']?.toString(),
          orElse: () => result,
        );
        await _showPurchaseReceiptOptions(purchaseRecord);
      }
      if (mounted) {
        if (!purchase &&
            result['autoReminders'] == true &&
            !remindersScheduled) {
          _notice(
              'Venta guardada con sus cuotas, pero no se pudo activar el aviso automático. Revisa la conexión del servidor.');
        } else if (!purchase && result['autoReminders'] == true) {
          _notice(
              'Venta registrada. Los avisos se programaron para el día antes de cada cuota.');
        } else {
          _notice(purchase
              ? 'Compra registrada; inventario actualizado.'
              : 'Venta registrada; inventario actualizado.');
        }
      }
    } catch (e) {
      if (mounted) {
        final message = '$e'.replaceFirst('Exception: ', '');
        _notice(!purchase && result['paymentMethod'] == 'pocket'
            ? 'Venta no registrada: $message'
            : message);
      }
    }
  }

  Future<void> _createAndShowSalePaylink(Map<String, dynamic> saleDraft) async {
    const configuredApi = String.fromEnvironment('BOT_API_URL');
    final root = _resolveApiRoot(configuredApi);
    if (root.isEmpty) throw Exception('No se encontró el servidor de cobros.');
    final phone = (saleDraft['reminderPhone']?.toString() ?? '')
        .replaceAll(RegExp(r'\D'), '');
    final amount = _number(saleDraft['pocketAmount']);
    if (phone.length < 7 || phone.length > 15) {
      throw Exception('Escribe un teléfono válido con código de país.');
    }
    if (amount <= 0) throw Exception('El total de la venta no es válido.');

    final response = await http
        .post(Uri.parse('$root/api/poket/paylinks'),
            headers: await _authorizedApiHeaders(json: true),
            body: jsonEncode({
              'saleId': saleDraft['id']?.toString() ?? '',
              'saleNumber': saleDraft['number']?.toString() ?? '',
              'customerName': saleDraft['party']?.toString() ?? 'cliente',
              'description':
                  'Venta ${saleDraft['number'] ?? ''} · ${saleDraft['party'] ?? 'cliente'}',
              'amount': amount,
              'currency': _store.currencyCode,
              'phone': phone,
              'sendWhatsApp': true,
              'saleDraft': saleDraft,
            }))
        .timeout(const Duration(seconds: 25));
    final body = jsonDecode(response.body) as Map<String, dynamic>;
    if (response.statusCode >= 300) {
      throw Exception(body['error'] ?? 'No se pudo crear el enlace de Poket.');
    }
    if (!mounted) return;
    final link = body['link']?.toString() ?? '';
    final whatsappMessage = body['sentWhatsApp'] == true
        ? 'El enlace se envió al cliente por WhatsApp.'
        : (body['whatsappError']?.toString().isNotEmpty == true
            ? body['whatsappError'].toString()
            : 'El enlace se creó; compártelo con el cliente.');
    var dialogOpen = true;
    var checking = false;
    var statusText =
        '$whatsappMessage La venta y el inventario no se guardan hasta que Poket confirme el pago.';
    StateSetter? updateDialog;
    BuildContext? paymentDialogContext;
    Timer? timer;
    Future<void> checkPayment() async {
      if (!dialogOpen || checking) return;
      checking = true;
      updateDialog?.call(
          () => statusText = 'Consultando el estado directamente en Poket…');
      try {
        final check = await http
            .get(
                Uri.parse(
                    '$root/api/poket/paylinks/${Uri.encodeComponent(body['id'].toString())}/status'),
                headers: await _authorizedApiHeaders())
            .timeout(const Duration(seconds: 18));
        final checked = jsonDecode(check.body) as Map<String, dynamic>;
        if (check.statusCode >= 300) {
          throw Exception(checked['error'] ?? 'No se pudo consultar el pago.');
        }
        final status = checked['status']?.toString() ?? 'pending';
        if (checked['confirmed'] == true && status == 'authorized') {
          final storedDraft = checked['saleDraft'] is Map
              ? Map<String, dynamic>.from(checked['saleDraft'] as Map)
              : saleDraft;
          await _store.recordTransaction(
            storedDraft,
            purchase: false,
            pocketPaymentConfirmed: true,
          );
          final sale = _store.sales.firstWhere(
            (item) => item['id']?.toString() == storedDraft['id']?.toString(),
            orElse: () => storedDraft,
          );
          dialogOpen = false;
          timer?.cancel();
          if (paymentDialogContext?.mounted == true) {
            Navigator.of(paymentDialogContext!).pop();
          }
          if (mounted) {
            _notice(
                'Poket confirmó el pago. Venta guardada y existencias actualizadas.');
            await _showReceiptOptions(sale);
          }
        } else if (status == 'expired' ||
            status == 'cancelled' ||
            status == 'failed') {
          updateDialog?.call(() => statusText =
              'El enlace terminó como ${status == 'expired' ? 'vencido' : status == 'cancelled' ? 'cancelado' : 'rechazado'}. No se guardó la venta ni se modificó el inventario.');
        } else {
          updateDialog?.call(() => statusText = checked['message']
                  ?.toString() ??
              'Poket todavía no confirma el pago. La venta sigue sin registrarse.');
        }
      } catch (error) {
        updateDialog
            ?.call(() => statusText = '$error'.replaceFirst('Exception: ', ''));
      } finally {
        checking = false;
      }
    }

    final autoPollUntil = DateTime.now().add(const Duration(minutes: 5));
    timer = Timer.periodic(const Duration(seconds: 7), (_) {
      if (DateTime.now().isBefore(autoPollUntil)) unawaited(checkPayment());
    });
    try {
      await showDialog<void>(
        context: context,
        barrierDismissible: false,
        builder: (dialogContext) => StatefulBuilder(
          builder: (dialogContext, setDialogState) {
            paymentDialogContext = dialogContext;
            updateDialog = setDialogState;
            return AlertDialog(
              title: const Text('Esperando confirmación de Poket'),
              content: SizedBox(
                width: 460,
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    SelectableText(link),
                    const SizedBox(height: 12),
                    Row(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          const Padding(
                            padding: EdgeInsets.only(right: 10, top: 2),
                            child: SizedBox(
                                width: 18,
                                height: 18,
                                child:
                                    CircularProgressIndicator(strokeWidth: 2)),
                          ),
                          Expanded(child: Text(statusText)),
                        ]),
                  ],
                ),
              ),
              actions: [
                TextButton(
                    onPressed: () => Navigator.pop(dialogContext),
                    child: const Text('Dejar pendiente')),
                FilledButton.icon(
                  onPressed: checking ? null : checkPayment,
                  icon: const Icon(Icons.refresh_rounded),
                  label: const Text('Comprobar pago'),
                ),
              ],
            );
          },
        ),
      );
    } finally {
      dialogOpen = false;
      timer.cancel();
    }
  }

  Future<void> _reconcilePendingPocketPayments() async {
    if (widget.demo || !mounted) return;
    const configuredApi = String.fromEnvironment('BOT_API_URL');
    final root = _resolveApiRoot(configuredApi);
    if (root.isEmpty) return;
    try {
      final response = await http
          .get(Uri.parse('$root/api/poket/paylinks/pending'),
              headers: await _authorizedApiHeaders())
          .timeout(const Duration(seconds: 12));
      if (response.statusCode >= 300) return;
      final body = jsonDecode(response.body) as Map<String, dynamic>;
      final pending =
          body['pending'] is List ? body['pending'] as List : const [];
      for (final entry in pending.whereType<Map>()) {
        if (!mounted) return;
        final id = entry['id']?.toString() ?? '';
        if (id.isEmpty) continue;
        final check = await http
            .get(
                Uri.parse(
                    '$root/api/poket/paylinks/${Uri.encodeComponent(id)}/status'),
                headers: await _authorizedApiHeaders())
            .timeout(const Duration(seconds: 18));
        if (check.statusCode >= 300) continue;
        final checked = jsonDecode(check.body) as Map<String, dynamic>;
        if (checked['confirmed'] != true || checked['saleDraft'] is! Map)
          continue;
        final draft = Map<String, dynamic>.from(checked['saleDraft'] as Map);
        await _store.recordTransaction(draft,
            purchase: false, pocketPaymentConfirmed: true);
        final sale = _store.sales.firstWhere(
          (item) => item['id']?.toString() == draft['id']?.toString(),
          orElse: () => draft,
        );
        if (mounted) {
          _notice(
              'Se recuperó una venta que Poket confirmó mientras no estabas en la app.');
          await _showReceiptOptions(sale);
        }
      }
    } catch (error) {
      debugPrint(
          'No se pudieron reconciliar cobros pendientes de Poket: $error');
    }
  }

  Future<void> _showReceiptOptions(Map<String, dynamic> sale) async {
    if (!mounted) return;
    final receipt = _store.receiptConfig;
    final catalog = _store.catalogConfig;
    final customer = _store.contacts.where((contact) =>
        contact['id']?.toString() == sale['customerId']?.toString() ||
        (sale['party']?.toString().trim().isNotEmpty == true &&
            contact['name']?.toString().trim().toLowerCase() ==
                sale['party']?.toString().trim().toLowerCase()));
    final knownCustomer =
        customer.isEmpty ? <String, dynamic>{} : customer.first;
    final phone = TextEditingController(
      text: (sale['reminderPhone']?.toString().isNotEmpty == true
                  ? sale['reminderPhone']
                  : knownCustomer['phone'])
              ?.toString() ??
          '',
    );
    final receiptKey = GlobalKey();
    final businessName =
        (receipt['businessName']?.toString().trim().isNotEmpty == true
                    ? receipt['businessName']
                    : catalog['name'])
                ?.toString() ??
            'Mi negocio';
    final logo = (receipt['logo']?.toString().isNotEmpty == true
                ? receipt['logo']
                : catalog['logo'])
            ?.toString() ??
        '';
    var busy = false;
    var errorText = '';

    Future<Uint8List> captureReceipt() async {
      await WidgetsBinding.instance.endOfFrame;
      final renderObject = receiptKey.currentContext?.findRenderObject();
      if (renderObject is! RenderRepaintBoundary) {
        throw Exception('No se pudo preparar la imagen del comprobante.');
      }
      final image = await renderObject.toImage(pixelRatio: 2);
      try {
        final data = await image.toByteData(format: ui.ImageByteFormat.png);
        if (data == null)
          throw Exception('No se pudo crear la imagen del comprobante.');
        return data.buffer.asUint8List();
      } finally {
        image.dispose();
      }
    }

    await showDialog<void>(
      context: context,
      builder: (dialogContext) => StatefulBuilder(
        builder: (dialogContext, setDialogState) => AlertDialog(
          title: const Text('Comprobante de venta'),
          content: SizedBox(
            width: 490,
            child: SingleChildScrollView(
              child: Column(
                mainAxisSize: MainAxisSize.min,
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  RepaintBoundary(
                    key: receiptKey,
                    child: _ReceiptCard(
                      sale: sale,
                      businessName: businessName,
                      logo: logo,
                      taxId: receipt['taxId']?.toString() ?? '',
                      address: receipt['address']?.toString() ?? '',
                      phone: receipt['phone']?.toString() ?? '',
                      email: receipt['email']?.toString() ?? '',
                    ),
                  ),
                  const SizedBox(height: 12),
                  TextField(
                    controller: phone,
                    keyboardType: TextInputType.phone,
                    decoration: const InputDecoration(
                      labelText: 'WhatsApp para enviar la imagen',
                      hintText: 'Código de país y número',
                      prefixIcon: Icon(Icons.phone_outlined),
                    ),
                  ),
                  if (errorText.isNotEmpty) ...[
                    const SizedBox(height: 8),
                    Text(errorText, style: const TextStyle(color: Colors.red)),
                  ],
                ],
              ),
            ),
          ),
          actions: [
            TextButton(
              onPressed: busy ? null : () => Navigator.pop(dialogContext),
              child: const Text('Cerrar'),
            ),
            OutlinedButton.icon(
              onPressed: busy
                  ? null
                  : () async {
                      setDialogState(() {
                        busy = true;
                        errorText = '';
                      });
                      try {
                        final bytes = await captureReceipt();
                        final document = pw.Document();
                        final receiptImage = pw.MemoryImage(bytes);
                        document.addPage(pw.Page(
                          pageFormat: PdfPageFormat.roll80,
                          margin: const pw.EdgeInsets.all(8),
                          build: (_) => pw.Center(
                            child:
                                pw.Image(receiptImage, fit: pw.BoxFit.contain),
                          ),
                        ));
                        await Printing.layoutPdf(
                          name:
                              'Comprobante-${sale['number'] ?? sale['id']}.pdf',
                          onLayout: (_) async => document.save(),
                        );
                        if (dialogContext.mounted) Navigator.pop(dialogContext);
                      } catch (error) {
                        setDialogState(() => errorText =
                            '$error'.replaceFirst('Exception: ', ''));
                      } finally {
                        if (dialogContext.mounted)
                          setDialogState(() => busy = false);
                      }
                    },
              icon: const Icon(Icons.print_outlined),
              label: const Text('Imprimir'),
            ),
            FilledButton.icon(
              onPressed: busy
                  ? null
                  : () async {
                      final digits = phone.text.replaceAll(RegExp(r'\D'), '');
                      if (digits.length < 7 || digits.length > 15) {
                        setDialogState(() => errorText =
                            'Escribe un número válido con código de país.');
                        return;
                      }
                      setDialogState(() {
                        busy = true;
                        errorText = '';
                      });
                      try {
                        const configuredApi =
                            String.fromEnvironment('BOT_API_URL');
                        final root = _resolveApiRoot(configuredApi);
                        if (root.isEmpty)
                          throw Exception(
                              'No se encontró el servidor de WhatsApp.');
                        final bytes = await captureReceipt();
                        final response = await http
                            .post(Uri.parse('$root/api/chats/receipt-image'),
                                headers:
                                    await _authorizedApiHeaders(json: true),
                                body: jsonEncode({
                                  'phone': digits,
                                  'customerName':
                                      sale['party']?.toString() ?? '',
                                  'caption':
                                      '$businessName · Comprobante ${sale['number'] ?? ''} · Total ${_money(_number(sale['total']))}',
                                  'image':
                                      'data:image/png;base64,${base64Encode(bytes)}',
                                }))
                            .timeout(const Duration(seconds: 25));
                        final result =
                            jsonDecode(response.body) as Map<String, dynamic>;
                        if (response.statusCode >= 300) {
                          throw Exception(result['error'] ??
                              'No se pudo enviar la imagen.');
                        }
                        if (dialogContext.mounted) Navigator.pop(dialogContext);
                        if (mounted)
                          _notice('Comprobante enviado por WhatsApp.');
                      } catch (error) {
                        setDialogState(() => errorText =
                            '$error'.replaceFirst('Exception: ', ''));
                      } finally {
                        if (dialogContext.mounted)
                          setDialogState(() => busy = false);
                      }
                    },
              icon: Icon(busy ? Icons.hourglass_top : Icons.send_rounded),
              label: Text(busy ? 'Procesando…' : 'Enviar WhatsApp'),
            ),
          ],
        ),
      ),
    );
    phone.dispose();
  }

  Future<void> _showPurchaseReceiptOptions(
      Map<String, dynamic> purchase) async {
    if (!mounted) return;
    final receipt = _store.receiptConfig;
    final catalog = _store.catalogConfig;
    final businessName =
        (receipt['businessName']?.toString().trim().isNotEmpty == true
                    ? receipt['businessName']
                    : catalog['name'])
                ?.toString() ??
            'Mi negocio';
    final logo = (receipt['logo']?.toString().isNotEmpty == true
                ? receipt['logo']
                : catalog['logo'])
            ?.toString() ??
        '';
    final receiptKey = GlobalKey();
    var busy = false;
    var errorText = '';

    Future<Uint8List> captureReceipt() async {
      await WidgetsBinding.instance.endOfFrame;
      final renderObject = receiptKey.currentContext?.findRenderObject();
      if (renderObject is! RenderRepaintBoundary) {
        throw Exception('No se pudo preparar el comprobante de compra.');
      }
      final image = await renderObject.toImage(pixelRatio: 2);
      try {
        final data = await image.toByteData(format: ui.ImageByteFormat.png);
        if (data == null) {
          throw Exception('No se pudo crear la imagen del comprobante.');
        }
        return data.buffer.asUint8List();
      } finally {
        image.dispose();
      }
    }

    await showDialog<void>(
      context: context,
      builder: (dialogContext) => StatefulBuilder(
        builder: (dialogContext, setDialogState) => AlertDialog(
          title: const Text('Comprobante de compra'),
          content: SizedBox(
            width: 430,
            child: SingleChildScrollView(
              child: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  RepaintBoundary(
                    key: receiptKey,
                    child: _PurchaseReceiptCard(
                      purchase: purchase,
                      businessName: businessName,
                      logo: logo,
                      taxId: receipt['taxId']?.toString() ?? '',
                      address: receipt['address']?.toString() ?? '',
                      phone: receipt['phone']?.toString() ?? '',
                      email: receipt['email']?.toString() ?? '',
                    ),
                  ),
                  if (errorText.isNotEmpty) ...[
                    const SizedBox(height: 10),
                    Text(errorText, style: const TextStyle(color: Colors.red)),
                  ],
                ],
              ),
            ),
          ),
          actions: [
            TextButton(
              onPressed: busy ? null : () => Navigator.pop(dialogContext),
              child: const Text('Cerrar'),
            ),
            FilledButton.icon(
              onPressed: busy
                  ? null
                  : () async {
                      setDialogState(() {
                        busy = true;
                        errorText = '';
                      });
                      try {
                        final bytes = await captureReceipt();
                        final document = pw.Document();
                        document.addPage(pw.Page(
                          pageFormat: PdfPageFormat.roll80,
                          margin: const pw.EdgeInsets.all(8),
                          build: (_) => pw.Center(
                            child: pw.Image(pw.MemoryImage(bytes),
                                fit: pw.BoxFit.contain),
                          ),
                        ));
                        await Printing.layoutPdf(
                          name:
                              'Comprobante-compra-${purchase['number'] ?? purchase['id']}.pdf',
                          onLayout: (_) async => document.save(),
                        );
                        if (dialogContext.mounted) {
                          Navigator.pop(dialogContext);
                        }
                      } catch (error) {
                        setDialogState(() => errorText =
                            '$error'.replaceFirst('Exception: ', ''));
                      } finally {
                        if (dialogContext.mounted) {
                          setDialogState(() => busy = false);
                        }
                      }
                    },
              icon: Icon(busy ? Icons.hourglass_top : Icons.print_outlined),
              label: Text(busy ? 'Preparando…' : 'Imprimir'),
            ),
          ],
        ),
      ),
    );
  }

  Future<void> _expenseDialog() async {
    final result = await showDialog<Map<String, dynamic>>(
        context: context, builder: (_) => const _ExpenseDialog());
    if (result == null) return;
    try {
      await _store.addRecord('expenses', result);
      if (mounted) _notice('Gasto guardado en Firebase.');
    } catch (error) {
      if (mounted) _notice('$error'.replaceFirst('Exception: ', ''));
    }
  }

  void _notice(String text) =>
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(text)));
}

class _Dashboard extends StatelessWidget {
  const _Dashboard({required this.store, required this.onNavigate});
  final _Store store;
  final ValueChanged<int> onNavigate;
  @override
  Widget build(BuildContext context) {
    final revenue = store.sales
        .fold<double>(0, (sum, item) => sum + _number(item['total']));
    final purchases = store.purchases
        .fold<double>(0, (sum, item) => sum + _number(item['total']));
    final lowStock = store.products
        .where((p) => _number(p['stock']) <= _number(p['minStock']))
        .length;
    return ListView(padding: const EdgeInsets.all(24), children: [
      const _PageHeading('Resumen', 'Así va tu negocio hoy.'),
      const SizedBox(height: 20),
      LayoutBuilder(builder: (context, constraints) {
        final count = constraints.maxWidth > 1000
            ? 4
            : constraints.maxWidth > 570
                ? 2
                : 1;
        return GridView.count(
            shrinkWrap: true,
            physics: const NeverScrollableScrollPhysics(),
            crossAxisCount: count,
            childAspectRatio: 2.2,
            crossAxisSpacing: 12,
            mainAxisSpacing: 12,
            children: [
              _MetricCard('Ventas registradas', _money(revenue),
                  Icons.trending_up, const Color(0xffe6f4eb)),
              _MetricCard('Compras registradas', _money(purchases),
                  Icons.shopping_bag_outlined, const Color(0xffeff1ff)),
              _MetricCard('Clientes', '${store.contacts.length}',
                  Icons.people_outline, const Color(0xfffff2dc)),
              _MetricCard('Stock bajo', '$lowStock productos',
                  Icons.inventory_2_outlined, const Color(0xffffe9e7)),
            ]);
      }),
      const SizedBox(height: 22),
      LayoutBuilder(builder: (context, constraints) {
        final activity = _Panel(
            title: 'Actividad reciente',
            trailing: TextButton(
                onPressed: () => onNavigate(1),
                child: const Text('Ver ventas')),
            child: store.sales.isEmpty
                ? const _Empty(
                    icon: Icons.point_of_sale_outlined,
                    text: 'Registra tu primera venta para empezar.')
                : Column(
                    children: store.sales
                        .take(5)
                        .map((r) => _RecordTile(
                            record: r, icon: Icons.point_of_sale_outlined))
                        .toList()));
        final quick = _Panel(
            title: 'Accesos rápidos',
            child: Column(children: [
              _QuickAction(
                  icon: Icons.add_shopping_cart,
                  label: 'Registrar venta',
                  onTap: () => onNavigate(1)),
              _QuickAction(
                  icon: Icons.add_box_outlined,
                  label: 'Añadir producto',
                  onTap: () => onNavigate(3)),
              _QuickAction(
                  icon: Icons.person_add_alt_1,
                  label: 'Agregar cliente',
                  onTap: () => onNavigate(4)),
              const Divider(),
              Text(
                  '${store.products.length} productos · ${store.purchases.length} compras · ${store.expenses.length} gastos',
                  style: const TextStyle(color: Colors.black54)),
            ]));
        if (constraints.maxWidth < 760)
          return Column(
              children: [activity, const SizedBox(height: 14), quick]);
        return Row(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Expanded(flex: 3, child: activity),
          const SizedBox(width: 16),
          Expanded(flex: 2, child: quick)
        ]);
      }),
    ]);
  }
}

class _ProductsPage extends StatelessWidget {
  const _ProductsPage({required this.store});
  final _Store store;
  @override
  Widget build(BuildContext context) =>
      ListView(padding: const EdgeInsets.all(24), children: [
        const _PageHeading(
            'Inventario', 'Controla existencias, precios y alertas de stock.'),
        const SizedBox(height: 18),
        if (store.products.isEmpty)
          const _Empty(
              icon: Icons.inventory_2_outlined,
              text:
                  'Aún no hay productos. Usa “Nuevo producto” para crear el primero.'),
        ...store.products.map((p) => Card(
            elevation: 0,
            child: ListTile(
              leading: _productThumbnail(p),
              title: Text(
                  '${p['name']}  ${_number(p['stock']) <= _number(p['minStock']) ? '· Stock bajo' : ''}',
                  style: const TextStyle(fontWeight: FontWeight.w700)),
              subtitle: Text(
                  '${p['sku'] ?? 'Sin código'} · Existencias: ${_number(p['stock'])} · Mínimo: ${_number(p['minStock'])}'),
              trailing: Row(mainAxisSize: MainAxisSize.min, children: [
                Text(_money(_number(p['price'])),
                    style: const TextStyle(fontWeight: FontWeight.w800)),
                IconButton(
                    icon: const Icon(Icons.edit_outlined),
                    onPressed: () => (context
                        .findAncestorStateOfType<_BusinessShellState>()
                        ?._productDialog(p))),
              ]),
            ))),
      ]);
}

class _ContactsPage extends StatelessWidget {
  const _ContactsPage({required this.store});
  final _Store store;
  @override
  Widget build(BuildContext context) {
    final debts = <String, List<Map<String, dynamic>>>{};
    for (final sale
        in store.sales.where((sale) => _number(sale['balance']) > 0)) {
      final linkedId = sale['customerId']?.toString();
      final matchingContact = store.contacts
          .where((contact) =>
              (linkedId != null && contact['id']?.toString() == linkedId) ||
              (linkedId == null &&
                  (contact['name']?.toString().trim().toLowerCase() ?? '') ==
                      (sale['party']?.toString().trim().toLowerCase() ?? '')))
          .firstOrNull;
      final id = matchingContact?['id']?.toString();
      if (id != null) debts.putIfAbsent(id, () => []).add(sale);
    }
    final totalDue = debts.values
        .expand((sales) => sales)
        .fold<double>(0, (sum, sale) => sum + _number(sale['balance']));
    return ListView(padding: const EdgeInsets.all(24), children: [
      const _PageHeading(
          'Clientes', 'Guarda los datos y el seguimiento de tus clientes.'),
      const SizedBox(height: 18),
      if (totalDue > 0)
        Card(
          elevation: 0,
          color: const Color(0xffedf3ff),
          child: ListTile(
            leading: const Icon(Icons.account_balance_wallet_outlined,
                color: _brand),
            title: const Text('Créditos pendientes',
                style: TextStyle(fontWeight: FontWeight.w800)),
            subtitle: Text('${debts.length} clientes con saldo pendiente'),
            trailing: Text(_money(totalDue),
                style: const TextStyle(
                    fontWeight: FontWeight.w800, color: _brand)),
          ),
        ),
      const SizedBox(height: 8),
      if (store.contacts.isEmpty)
        const _Empty(
            icon: Icons.people_outline,
            text:
                'Aún no hay clientes. Agrega el primero con “Nuevo cliente”.'),
      ...store.contacts.map((contact) {
        final contactSales =
            debts[contact['id']?.toString()] ?? const <Map<String, dynamic>>[];
        final due = contactSales.fold<double>(
            0, (sum, sale) => sum + _number(sale['balance']));
        return Card(
          elevation: 0,
          child: Padding(
            padding: const EdgeInsets.symmetric(vertical: 4),
            child: Column(children: [
              ListTile(
                leading: CircleAvatar(
                  backgroundColor: const Color(0xffedf3ff),
                  child: Text(
                    String.fromCharCode(
                            (contact['name'] as String? ?? '?').runes.first)
                        .toUpperCase(),
                    style: const TextStyle(color: _brand),
                  ),
                ),
                title: Text(contact['name'] ?? '',
                    style: const TextStyle(fontWeight: FontWeight.w700)),
                subtitle: Text([
                  contact['phone'],
                  contact['email'],
                  contact['notes']
                ]
                    .where(
                        (value) => value != null && value.toString().isNotEmpty)
                    .join(' · ')),
                trailing: IconButton(
                  tooltip: 'Editar cliente',
                  icon: const Icon(Icons.edit_outlined),
                  onPressed: () => context
                      .findAncestorStateOfType<_BusinessShellState>()
                      ?._contactDialog(contact),
                ),
              ),
              if (due > 0)
                Padding(
                  padding: const EdgeInsets.fromLTRB(72, 0, 16, 12),
                  child: Row(children: [
                    Expanded(
                        child: Text(
                            '${contactSales.length} crédito(s) · Debe ${_money(due)}',
                            style: const TextStyle(
                                fontWeight: FontWeight.w700, color: _brand))),
                    IconButton.filledTonal(
                      tooltip: 'Enviar recordatorio por correo',
                      onPressed:
                          contact['email']?.toString().trim().isNotEmpty == true
                              ? () => _sendCreditReminder(context, store,
                                  contact, contactSales, 'email')
                              : null,
                      icon: const Icon(Icons.email_outlined),
                    ),
                    const SizedBox(width: 6),
                    IconButton.filledTonal(
                      tooltip: 'Enviar recordatorio por WhatsApp',
                      onPressed:
                          contact['phone']?.toString().trim().isNotEmpty == true
                              ? () => _sendCreditReminder(context, store,
                                  contact, contactSales, 'whatsapp')
                              : null,
                      icon: const Icon(Icons.chat_outlined),
                    ),
                  ]),
                ),
            ]),
          ),
        );
      }),
    ]);
  }
}

Future<void> _sendCreditReminder(
  BuildContext context,
  _Store store,
  Map<String, dynamic> contact,
  List<Map<String, dynamic>> sales,
  String channel,
) async {
  final name = contact['name']?.toString() ?? 'cliente';
  final amount =
      sales.fold<double>(0, (sum, sale) => sum + _number(sale['balance']));
  final currency = _money(amount);
  final invoices = sales
      .map((sale) =>
          '${sale['number'] ?? 'Venta'}: ${_money(_number(sale['balance']))}')
      .join('\n');
  final message =
      'Hola $name, te compartimos un recordatorio amable sobre tu saldo pendiente de $currency.\n$invoices\nSi ya realizaste el pago, ignora este mensaje. ¡Gracias!';
  try {
    if (channel == 'email') {
      final email = contact['email']?.toString().trim() ?? '';
      final query = {
        'subject': 'Recordatorio de saldo pendiente',
        'body': message
      }
          .entries
          .map((entry) =>
              '${Uri.encodeComponent(entry.key)}=${Uri.encodeComponent(entry.value)}')
          .join('&');
      final uri = Uri.parse('mailto:${Uri.encodeComponent(email)}?$query');
      if (!await launchUrl(uri))
        throw Exception('No encontramos una aplicación de correo configurada.');
      if (context.mounted)
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
              content: Text(
                  'Se abrió el correo con el recordatorio; revísalo y pulsa Enviar.')),
        );
    } else {
      const configuredApi = String.fromEnvironment('BOT_API_URL');
      final root = _resolveApiRoot(configuredApi);
      if (root.isEmpty)
        throw Exception(
            'Configura BOT_API_URL para enviar por WhatsApp desde esta plataforma.');
      final uri = Uri.parse('$root/api/reminders/whatsapp');
      final response = await http
          .post(
            uri,
            headers: await _authorizedApiHeaders(json: true),
            body: jsonEncode({'phone': contact['phone'], 'message': message}),
          )
          .timeout(const Duration(seconds: 20));
      final data = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300)
        throw Exception(data['error'] ?? 'No se pudo enviar el recordatorio.');
      if (context.mounted)
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Recordatorio enviado por WhatsApp.')),
        );
    }
  } catch (error) {
    if (context.mounted)
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('$error'.replaceFirst('Exception: ', ''))),
      );
  }
}

class _RecordsPage extends StatelessWidget {
  const _RecordsPage(
      {required this.title,
      required this.subtitle,
      required this.icon,
      required this.records,
      required this.money,
      this.store});
  final String title, subtitle;
  final IconData icon;
  final List<Map<String, dynamic>> records;
  final bool money;
  final _Store? store;
  @override
  Widget build(BuildContext context) =>
      ListView(padding: const EdgeInsets.all(24), children: [
        _PageHeading(title, subtitle),
        const SizedBox(height: 18),
        if (records.isEmpty)
          _Empty(
              icon: icon,
              text:
                  'Todavía no hay registros de $title. Usa el botón para agregar uno.'),
        ...records.map((r) => _RecordTile(
              record: r,
              icon: icon,
              money: money,
              onPayment: store == null || _number(r['balance']) <= 0
                  ? null
                  : () => _paymentDialog(context, store!, r),
              onPoketPayment: title != 'Ventas' ||
                      store == null ||
                      store!.demo ||
                      _number(r['balance']) <= 0
                  ? null
                  : () => _poketPaymentDialog(context, store!, r),
              onReceipt: title == 'Ventas'
                  ? () => context
                      .findAncestorStateOfType<_BusinessShellState>()
                      ?._showReceiptOptions(r)
                  : title == 'Compras'
                      ? () => context
                          .findAncestorStateOfType<_BusinessShellState>()
                          ?._showPurchaseReceiptOptions(r)
                      : null,
              onReturn: title != 'Ventas' || store == null || !_canReturn(r)
                  ? null
                  : () => _returnDialog(context, store!, r),
            )),
        if (title == 'Ventas' && store?.returns.isNotEmpty == true) ...[
          const SizedBox(height: 18),
          const _PageHeading('Devoluciones recientes',
              'Historial de productos devueltos y ajustes registrados.'),
          const SizedBox(height: 8),
          ...store!.returns.take(20).map((returnRecord) {
            final returnedItems = returnRecord['items'] is List
                ? (returnRecord['items'] as List)
                    .whereType<Map>()
                    .map((item) => '${item['quantity']} × ${item['name']}')
                    .join(', ')
                : '';
            return Card(
              elevation: 0,
              child: ListTile(
                leading: const CircleAvatar(
                  backgroundColor: Color(0xfffff2dc),
                  child: Icon(Icons.assignment_return_outlined, color: _brand),
                ),
                title: Text(
                    'Venta ${returnRecord['saleNumber'] ?? ''} · ${_money(_number(returnRecord['total']))}'),
                subtitle: Text(
                    '${returnRecord['date'] ?? ''} · ${returnRecord['customer'] ?? ''}\n$returnedItems'),
                isThreeLine: true,
              ),
            );
          }),
        ],
      ]);

  bool _canReturn(Map<String, dynamic> sale) {
    final rawItems = sale['items'];
    if (rawItems is! List) return false;
    final returnedRaw = sale['returnedQuantities'];
    final returned = returnedRaw is Map
        ? returnedRaw
            .map((key, value) => MapEntry(key.toString(), _number(value)))
        : <String, double>{};
    return rawItems.whereType<Map>().any((item) {
      final id = item['productId']?.toString() ?? '';
      return _number(item['quantity']) > (returned[id] ?? 0) + .000001;
    });
  }

  Future<void> _returnDialog(
      BuildContext context, _Store store, Map<String, dynamic> sale) async {
    final result = await showDialog<Map<String, dynamic>>(
      context: context,
      builder: (_) => _SaleReturnDialog(sale: sale),
    );
    if (result == null || !context.mounted) return;
    try {
      final saved = await store.recordSaleReturn(
        saleId: sale['id']?.toString() ?? '',
        items: (result['items'] as List).cast<Map<String, dynamic>>(),
        reason: result['reason']?.toString() ?? '',
        refundMethod: result['refundMethod']?.toString() ?? '',
      );
      if (context.mounted) {
        final refund = _number(saved['refundAmount']);
        final offset = _number(saved['balanceAdjusted']);
        final message = refund > 0
            ? 'Devolución registrada. Reembolso de ${_money(refund)} anotado por ${saved['refundMethod']}; inventario actualizado.'
            : 'Devolución registrada. Se ajustaron ${_money(offset)} del saldo pendiente y se actualizó el inventario.';
        ScaffoldMessenger.of(context)
            .showSnackBar(SnackBar(content: Text(message)));
      }
    } catch (error) {
      if (context.mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('$error'.replaceFirst('Exception: ', ''))),
        );
      }
    }
  }

  Future<void> _poketPaymentDialog(
      BuildContext context, _Store store, Map<String, dynamic> sale) async {
    const configuredUrl = String.fromEnvironment('BOT_API_URL');
    final root = _resolveApiRoot(configuredUrl);
    if (root.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(const SnackBar(
          content: Text('No se encontró el servidor para generar el cobro.')));
      return;
    }
    final balance = _number(sale['balance']);
    final amount = TextEditingController(text: balance.toStringAsFixed(2));
    final phone =
        TextEditingController(text: (sale['reminderPhone'] ?? '').toString());
    final digits = phone.text.replaceAll(RegExp(r'\D'), '');
    var sendWhatsApp = digits.length >= 7 && digits.length <= 15;
    var saving = false;
    String? errorText;
    final result = await showDialog<Map<String, dynamic>>(
      context: context,
      builder: (dialogContext) => StatefulBuilder(
        builder: (dialogContext, refresh) => AlertDialog(
          title: const Text('Cobrar con Poket'),
          content: SizedBox(
            width: 390,
            child: Column(mainAxisSize: MainAxisSize.min, children: [
              Text('Saldo pendiente: ${_money(balance)}'),
              const SizedBox(height: 12),
              TextField(
                controller: amount,
                keyboardType:
                    const TextInputType.numberWithOptions(decimal: true),
                decoration: const InputDecoration(labelText: 'Monto del cobro'),
              ),
              const SizedBox(height: 10),
              TextField(
                controller: phone,
                keyboardType: TextInputType.phone,
                decoration: const InputDecoration(
                    labelText: 'WhatsApp del cliente',
                    helperText: 'Incluye el código de país.'),
              ),
              CheckboxListTile(
                contentPadding: EdgeInsets.zero,
                value: sendWhatsApp,
                title: const Text('Enviar el enlace por WhatsApp'),
                onChanged: saving
                    ? null
                    : (value) => refresh(() => sendWhatsApp = value ?? false),
              ),
              if (errorText != null)
                Align(
                    alignment: Alignment.centerLeft,
                    child: Text(errorText!,
                        style: const TextStyle(color: Colors.red))),
            ]),
          ),
          actions: [
            TextButton(
                onPressed: saving ? null : () => Navigator.pop(dialogContext),
                child: const Text('Cancelar')),
            FilledButton.icon(
              onPressed: saving
                  ? null
                  : () async {
                      final charge = _number(amount.text);
                      final customerPhone =
                          phone.text.replaceAll(RegExp(r'\D'), '');
                      if (charge <= 0 || charge > balance) {
                        refresh(() => errorText =
                            'El monto debe ser mayor que cero y no superar el saldo pendiente.');
                        return;
                      }
                      if (sendWhatsApp &&
                          (customerPhone.length < 7 ||
                              customerPhone.length > 15)) {
                        refresh(() => errorText =
                            'Escribe un número válido con código de país.');
                        return;
                      }
                      refresh(() {
                        saving = true;
                        errorText = null;
                      });
                      try {
                        final response = await http
                            .post(Uri.parse('$root/api/poket/paylinks'),
                                headers:
                                    await _authorizedApiHeaders(json: true),
                                body: jsonEncode({
                                  'saleId': sale['id']?.toString() ?? '',
                                  'saleNumber':
                                      sale['number']?.toString() ?? '',
                                  'customerName':
                                      sale['party']?.toString() ?? 'cliente',
                                  'description':
                                      'Venta ${sale['number'] ?? ''} · ${sale['party'] ?? 'cliente'}',
                                  'amount': charge,
                                  'currency': sale['currency']?.toString() ??
                                      store.currencyCode,
                                  'phone': customerPhone,
                                  'sendWhatsApp': sendWhatsApp,
                                }))
                            .timeout(const Duration(seconds: 25));
                        final body =
                            jsonDecode(response.body) as Map<String, dynamic>;
                        if (response.statusCode >= 300) {
                          throw Exception(body['error'] ??
                              'No se pudo crear el enlace de Poket.');
                        }
                        if (dialogContext.mounted) {
                          Navigator.pop(dialogContext, body);
                        }
                      } catch (e) {
                        refresh(() {
                          saving = false;
                          errorText = '$e'.replaceFirst('Exception: ', '');
                        });
                      }
                    },
              icon: Icon(saving ? Icons.hourglass_top : Icons.link),
              label: Text(saving ? 'Creando enlace…' : 'Crear PayLink'),
            ),
          ],
        ),
      ),
    );
    amount.dispose();
    phone.dispose();
    if (result == null || !context.mounted) return;
    final link = result['link']?.toString() ?? '';
    await showDialog<void>(
      context: context,
      builder: (resultContext) => AlertDialog(
        title: const Text('Enlace de Poket creado'),
        content: SizedBox(
          width: 420,
          child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                SelectableText(link),
                const SizedBox(height: 12),
                Text(result['sentWhatsApp'] == true
                    ? 'El enlace se envió al cliente por WhatsApp.'
                    : (result['whatsappError']?.toString().isNotEmpty == true
                        ? result['whatsappError'].toString()
                        : 'Copia el enlace y compártelo con el cliente.')),
                const SizedBox(height: 6),
                const Text('Es de un solo uso y vence en 7 días.'),
              ]),
        ),
        actions: [
          TextButton.icon(
            onPressed: () async {
              await Clipboard.setData(ClipboardData(text: link));
              if (resultContext.mounted) Navigator.pop(resultContext);
              if (context.mounted) {
                ScaffoldMessenger.of(context).showSnackBar(const SnackBar(
                    content: Text('Enlace copiado al portapapeles.')));
              }
            },
            icon: const Icon(Icons.copy),
            label: const Text('Copiar'),
          ),
          TextButton(
              onPressed: () => Navigator.pop(resultContext),
              child: const Text('Cerrar')),
        ],
      ),
    );
  }

  Future<void> _paymentDialog(
      BuildContext context, _Store store, Map<String, dynamic> record) async {
    final controller =
        TextEditingController(text: '${_number(record['balance'])}');
    final amount = await showDialog<double>(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('Registrar abono'),
        content: SizedBox(
          width: 380,
          child: _field(controller, 'Monto recibido', number: true),
        ),
        actions: [
          TextButton(
              onPressed: () => Navigator.pop(context),
              child: const Text('Cancelar')),
          FilledButton(
            onPressed: () {
              final value = _number(controller.text);
              if (value > 0 && value <= _number(record['balance']))
                Navigator.pop(context, value);
            },
            child: const Text('Guardar abono'),
          ),
        ],
      ),
    );
    controller.dispose();
    if (amount == null) return;
    try {
      final remindersUpdated = await store.recordPayment(record, amount);
      if (context.mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
              content: Text(record['autoReminders'] == true && !remindersUpdated
                  ? 'Abono registrado, pero no se pudo actualizar la agenda de avisos.'
                  : 'Abono registrado.')),
        );
      }
    } catch (error) {
      if (context.mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('$error'.replaceFirst('Exception: ', ''))),
        );
      }
    }
  }
}

class _RecordTile extends StatelessWidget {
  const _RecordTile(
      {required this.record,
      required this.icon,
      this.money = true,
      this.onPayment,
      this.onPoketPayment,
      this.onReceipt,
      this.onReturn});
  final Map<String, dynamic> record;
  final IconData icon;
  final bool money;
  final VoidCallback? onPayment;
  final VoidCallback? onPoketPayment;
  final VoidCallback? onReceipt;
  final VoidCallback? onReturn;
  @override
  Widget build(BuildContext context) {
    final schedule = (record['paymentSchedule'] as List? ?? [])
        .whereType<Map>()
        .map((item) => Map<String, dynamic>.from(item))
        .toList();
    final pending = schedule
        .where((item) =>
            item['status'] != 'paid' &&
            _number(item['remainingAmount'] ?? item['amount']) > 0)
        .toList();
    final nextDue = pending.isEmpty
        ? null
        : DateTime.tryParse(pending.first['dueAt']?.toString() ?? '')
            ?.toLocal();
    final scheduleText = nextDue == null
        ? ''
        : ' · ${pending.length} ${pending.length == 1 ? 'cuota' : 'cuotas'} · Próxima: ${DateFormat('dd/MM/yyyy').format(nextDue)}';
    return Card(
        elevation: 0,
        child: ListTile(
          leading: CircleAvatar(
              backgroundColor: const Color(0xffedf3ff),
              child: Icon(icon, color: _brand)),
          title: Row(children: [
            Expanded(
                child: Text(record['label'] ?? record['name'] ?? 'Registro',
                    style: const TextStyle(fontWeight: FontWeight.w700))),
            if (onPoketPayment != null)
              IconButton(
                tooltip: 'Crear o enviar PayLink de Poket',
                visualDensity: VisualDensity.compact,
                onPressed: onPoketPayment,
                icon: const Icon(Icons.link, color: _brand),
              ),
            if (onReceipt != null)
              IconButton(
                tooltip: 'Ver o imprimir comprobante',
                visualDensity: VisualDensity.compact,
                onPressed: onReceipt,
                icon: const Icon(Icons.receipt_long_outlined, color: _brand),
              ),
            if (onReturn != null)
              IconButton(
                tooltip: 'Registrar devolución',
                visualDensity: VisualDensity.compact,
                onPressed: onReturn,
                icon:
                    const Icon(Icons.assignment_return_outlined, color: _brand),
              ),
          ]),
          subtitle: Text(
              '${record['number'] ?? ''} ${record['date'] ?? ''} ${record['party'] ?? ''} ${record['detail'] ?? ''}'
                      .trim() +
                  (_number(record['balance']) > 0
                      ? ' · Pendiente: ${_money(_number(record['balance']))}'
                      : '') +
                  scheduleText),
          trailing: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.end,
            children: [
              Text(_money(_number(record['total'])),
                  style: TextStyle(
                      fontWeight: FontWeight.w800,
                      color: _number(record['balance']) > 0 ? _brand : _ink)),
              if (_number(record['balance']) > 0)
                Text('Debe ${_money(_number(record['balance']))}',
                    style: const TextStyle(fontSize: 11, color: _brand)),
              if (onPayment != null)
                SizedBox(
                  height: 30,
                  child: TextButton(
                    onPressed: onPayment,
                    style: TextButton.styleFrom(
                        padding: const EdgeInsets.symmetric(horizontal: 8)),
                    child: const Text('Abonar', style: TextStyle(fontSize: 12)),
                  ),
                ),
            ],
          ),
        ));
  }
}

class _Store extends ChangeNotifier {
  _Store(this.uid, this.demo);
  final String uid;
  final bool demo;
  bool loading = true;
  bool refreshing = false;
  bool _loadRunning = false;
  String? loadWarning;
  String currencyCode = 'NIO';
  Map<String, dynamic> catalogConfig = {};
  Map<String, dynamic> receiptConfig = {};
  List<Map<String, dynamic>> products = [],
      contacts = [],
      sales = [],
      returns = [],
      purchases = [],
      expenses = [];
  void clearSessionView() {
    products = [];
    contacts = [];
    sales = [];
    returns = [];
    purchases = [];
    expenses = [];
    catalogConfig = {};
    receiptConfig = {};
    currencyCode = 'NIO';
    _setCurrentCurrency(currencyCode);
    loadWarning = null;
    refreshing = false;
    loading = false;
    notifyListeners();
  }

  static const _keys = [
    'products',
    'contacts',
    'sales',
    'returns',
    'purchases',
    'expenses'
  ];
  CollectionReference<Map<String, dynamic>> _collection(String name) =>
      FirebaseFirestore.instance.collection('users').doc(uid).collection(name);
  void _warmPaymentReminderWorker() {
    const configuredApi = String.fromEnvironment('BOT_API_URL');
    final root = _resolveApiRoot(configuredApi);
    if (demo || root.isEmpty) return;
    unawaited(() async {
      try {
        await http
            .get(Uri.parse('$root/api/state'),
                headers: await _authorizedApiHeaders())
            .timeout(const Duration(seconds: 12));
      } catch (error) {
        debugPrint(
            'No se pudo activar la sesión del servidor de avisos: $error');
      }
    }());
  }

  Future<void> load() async {
    if (_loadRunning) return;
    _loadRunning = true;
    loadWarning = null;
    refreshing = true;
    try {
      if (demo) {
        final prefs = await SharedPreferences.getInstance();
        currencyCode = prefs.getString('sync-connect:currency') ?? 'NIO';
        final catalogJson = prefs.getString('sync-connect:catalog');
        if (catalogJson != null) {
          catalogConfig = Map<String, dynamic>.from(jsonDecode(catalogJson));
        }
        final receiptJson = prefs.getString('sync-connect:receipt');
        if (receiptJson != null) {
          receiptConfig = Map<String, dynamic>.from(jsonDecode(receiptJson));
        }
        for (final key in _keys) {
          final raw = prefs.getString('sync-connect:$key');
          if (raw != null)
            _setList(
                key, (jsonDecode(raw) as List).cast<Map<String, dynamic>>());
        }
        if (products.isEmpty) {
          products = [
            {
              'id': 'seed-cafe',
              'name': 'Café de la casa',
              'sku': 'CAF-001',
              'price': 35.0,
              'cost': 18.0,
              'stock': 24.0,
              'minStock': 5.0
            },
            {
              'id': 'seed-pan',
              'name': 'Pan artesanal',
              'sku': 'PAN-002',
              'price': 12.0,
              'cost': 5.0,
              'stock': 3.0,
              'minStock': 5.0
            },
            {
              'id': 'seed-mug',
              'name': 'Taza reutilizable',
              'sku': 'ACC-003',
              'price': 55.0,
              'cost': 25.0,
              'stock': 12.0,
              'minStock': 3.0
            },
          ];
          contacts = [
            {
              'id': 'seed-customer',
              'name': 'María López',
              'phone': '5555-0101',
              'email': 'maria@example.com',
              'notes': 'Cliente frecuente'
            }
          ];
          sales = [
            {
              'id': 'seed-sale',
              'label': 'Venta de ejemplo',
              'number': 'V-0001',
              'date': DateFormat('dd/MM/yyyy').format(DateTime.now()),
              'total': 105.0,
              'detail': '3 productos'
            }
          ];
        }
      } else {
        final prefs = await SharedPreferences.getInstance();
        final cachedPreferences =
            prefs.getString('sync-connect:firebase:$uid:preferences');
        if (cachedPreferences != null) {
          final data = Map<String, dynamic>.from(jsonDecode(cachedPreferences));
          currencyCode = data['currency']?.toString() ?? currencyCode;
          if (data['catalog'] is Map)
            catalogConfig = Map<String, dynamic>.from(data['catalog'] as Map);
          if (data['receipt'] is Map)
            receiptConfig = Map<String, dynamic>.from(data['receipt'] as Map);
        }
        for (final key in _keys) {
          final cached = prefs.getString('sync-connect:firebase:$uid:$key');
          if (cached != null) {
            try {
              _setList(key, _toMapList(jsonDecode(cached)));
            } catch (_) {}
          }
        }
        _setCurrentCurrency(currencyCode);
        loading = false;
        notifyListeners();
        _warmPaymentReminderWorker();

        final tasks = <Future<Object?>>[
          for (final key in _keys)
            () async {
              try {
                final snapshot = await _collection(key)
                    .get()
                    .timeout(const Duration(seconds: 30));
                final values = snapshot.docs
                    .map((d) => {'id': d.id, ...d.data()})
                    .toList();
                _setList(key, values);
                await prefs.setString(
                    'sync-connect:firebase:$uid:$key', jsonEncode(values));
                return null;
              } catch (e) {
                debugPrint('Firestore ($key): $e');
                return e;
              }
            }(),
          () async {
            try {
              final settings = await _collection('settings')
                  .doc('preferences')
                  .get()
                  .timeout(const Duration(seconds: 30));
              final data = settings.data() ?? <String, dynamic>{};
              currencyCode = data['currency']?.toString() ?? 'NIO';
              if (data['catalog'] is Map)
                catalogConfig =
                    Map<String, dynamic>.from(data['catalog'] as Map);
              if (data['receipt'] is Map)
                receiptConfig =
                    Map<String, dynamic>.from(data['receipt'] as Map);
              _setCurrentCurrency(currencyCode);
              await prefs.setString(
                  'sync-connect:firebase:$uid:preferences', jsonEncode(data));
              return null;
            } catch (e) {
              debugPrint('Firestore (preferencias): $e');
              return e;
            }
          }(),
        ];
        final errors = (await Future.wait(tasks)).whereType<Object>().toList();
        if (errors.isNotEmpty)
          loadWarning = _firestoreLoadMessage(errors.first);
      }
      _setCurrentCurrency(currencyCode);
    } catch (e) {
      loadWarning =
          'No se pudieron cargar tus datos. Revisa la conexión y vuelve a intentarlo.';
      debugPrint('No se pudieron cargar los datos: $e');
    } finally {
      loading = false;
      refreshing = false;
      _loadRunning = false;
      notifyListeners();
    }
  }

  String _firestoreLoadMessage(Object error) {
    if (error is TimeoutException) {
      return 'Firestore está tardando demasiado en responder. Comprueba que la base de datos y sus reglas estén activas.';
    }
    if (error is FirebaseException) {
      switch (error.code) {
        case 'permission-denied':
          return 'Firebase denegó el acceso a tus datos. Publica las reglas de Firestore del proyecto.';
        case 'unauthenticated':
          return 'Tu sesión de Firebase venció. Cierra sesión e inicia sesión de nuevo.';
        case 'unavailable':
        case 'deadline-exceeded':
          return 'La sincronización está pendiente; la aplicación conserva los datos locales disponibles.';
        case 'failed-precondition':
          return 'Firestore aún no está configurado. Revisa que la base de datos esté creada en Firebase.';
        default:
          return 'Firestore no pudo actualizar tus datos (${error.code}). Vuelve a intentar.';
      }
    }
    return 'No se pudieron actualizar algunos datos. Vuelve a intentar.';
  }

  String _firestoreWriteMessage(Object error, String item) {
    if (error is FirebaseException) {
      switch (error.code) {
        case 'permission-denied':
          return 'Firebase rechazó guardar $item. Publica las reglas de firestore.rules en el proyecto cmr-boot e inicia sesión con tu cuenta.';
        case 'unauthenticated':
          return 'Tu sesión venció. Inicia sesión de nuevo para guardar $item.';
        case 'unavailable':
        case 'deadline-exceeded':
          return 'Firebase no está disponible ahora. $item no se guardó; vuelve a intentarlo.';
        default:
          return 'No se pudo guardar $item en Firebase (${error.code}). ${error.message ?? ''}';
      }
    }
    return 'No se pudo guardar $item en Firebase: $error';
  }

  Future<void> updateCurrency(String code) async {
    if (!_currencyNames.containsKey(code)) return;
    if (demo) {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString('sync-connect:currency', code);
    } else {
      try {
        await _collection('settings')
            .doc('preferences')
            .set({'currency': code}, SetOptions(merge: true));
      } catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'la moneda'));
      }
    }
    currencyCode = code;
    _setCurrentCurrency(code);
    notifyListeners();
  }

  Future<void> updateCatalogConfig(Map<String, dynamic> config) async {
    final nextConfig = Map<String, dynamic>.from(config);
    if (demo) {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString('sync-connect:catalog', jsonEncode(nextConfig));
    } else {
      try {
        await _collection('settings')
            .doc('preferences')
            .set({'catalog': nextConfig}, SetOptions(merge: true));
      } catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'el catálogo'));
      }
    }
    catalogConfig = nextConfig;
    notifyListeners();
  }

  Future<void> updateReceiptConfig(Map<String, dynamic> config) async {
    final nextConfig = Map<String, dynamic>.from(config);
    if (demo) {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString('sync-connect:receipt', jsonEncode(nextConfig));
    } else {
      try {
        await _collection('settings')
            .doc('preferences')
            .set({'receipt': nextConfig}, SetOptions(merge: true));
        final prefs = await SharedPreferences.getInstance();
        await prefs.setString(
            'sync-connect:firebase:$uid:preferences',
            jsonEncode({
              'currency': currencyCode,
              'catalog': catalogConfig,
              'receipt': nextConfig,
            }));
      } catch (error) {
        throw Exception(
            _firestoreWriteMessage(error, 'los datos de facturación'));
      }
    }
    receiptConfig = nextConfig;
    notifyListeners();
  }

  void _setList(String key, List<Map<String, dynamic>> value) {
    switch (key) {
      case 'products':
        products = value;
        break;
      case 'contacts':
        contacts = value;
        break;
      case 'sales':
        sales = value;
        break;
      case 'returns':
        returns = value;
        break;
      case 'purchases':
        purchases = value;
        break;
      case 'expenses':
        expenses = value;
        break;
    }
  }

  List<Map<String, dynamic>> _list(String key) => switch (key) {
        'products' => products,
        'contacts' => contacts,
        'sales' => sales,
        'returns' => returns,
        'purchases' => purchases,
        _ => expenses
      };
  Future<void> _persist(String key) async {
    if (demo) {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString('sync-connect:$key', jsonEncode(_list(key)));
    }
    notifyListeners();
  }

  Future<void> saveProduct(Map<String, dynamic> data, {String? id}) async {
    final newId = id ?? DateTime.now().microsecondsSinceEpoch.toString();
    final record = {'id': newId, ...data};
    if (!demo) {
      try {
        await _collection('products').doc(newId).set(data);
      } catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'el producto'));
      }
    }
    final index = products.indexWhere((p) => p['id'] == newId);
    if (index < 0) {
      products = [record, ...products];
    } else {
      products[index] = record;
    }
    await _persist('products');
  }

  Future<void> saveContact(Map<String, dynamic> data, {String? id}) async {
    final newId = id ?? DateTime.now().microsecondsSinceEpoch.toString();
    final record = {'id': newId, ...data};
    if (!demo) {
      try {
        await _collection('contacts').doc(newId).set(data);
      } catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'el cliente'));
      }
    }
    final index = contacts.indexWhere((p) => p['id'] == newId);
    if (index < 0) {
      contacts = [record, ...contacts];
    } else {
      contacts[index] = record;
    }
    await _persist('contacts');
  }

  Future<void> addRecord(String key, Map<String, dynamic> data) async {
    final id = DateTime.now().microsecondsSinceEpoch.toString();
    final record = {'id': id, ...data};
    if (!demo) {
      try {
        await _collection(key).doc(id).set(data);
      } catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'el registro'));
      }
    }
    _list(key).insert(0, record);
    await _persist(key);
  }

  Future<bool> recordTransaction(Map<String, dynamic> data,
      {required bool purchase, bool pocketPaymentConfirmed = false}) async {
    final collectionName = purchase ? 'purchases' : 'sales';
    if (!purchase &&
        data['paymentMethod'] == 'pocket' &&
        !pocketPaymentConfirmed) {
      throw Exception(
          'La venta de Pocket solo se registra después de confirmar el pago.');
    }
    final supplied = data['items'] is List
        ? (data['items'] as List)
            .whereType<Map>()
            .map((item) => Map<String, dynamic>.from(item))
            .toList()
        : [
            {
              'productId': data['productId'],
              'quantity': data['quantity'],
            }
          ];
    if (supplied.isEmpty) throw Exception('Agrega productos a la operación.');
    final requested = <String, double>{};
    for (final item in supplied) {
      final productId = item['productId']?.toString() ?? '';
      final qty = _number(item['quantity']);
      if (productId.isEmpty || qty <= 0) {
        throw Exception('Revisa los productos y sus cantidades.');
      }
      requested.update(productId, (current) => current + qty,
          ifAbsent: () => qty);
    }
    final id = data['id']?.toString() ??
        DateTime.now().microsecondsSinceEpoch.toString();
    final transactionCurrency =
        _currencyNames.containsKey(data['currency']?.toString())
            ? data['currency'].toString()
            : currencyCode;
    final lineItems = <Map<String, dynamic>>[];
    var paymentSchedule = <Map<String, dynamic>>[];
    var total = 0.0;
    var subtotal = 0.0;
    var pocketSurcharge = 0.0;
    var alreadyRecorded = false;
    final installmentCount =
        _number(data['installmentCount']).toInt().clamp(1, 12);
    final frequency = data['installmentFrequency']?.toString() ?? 'monthly';
    final firstDueDate =
        DateTime.tryParse(data['firstDueAt']?.toString() ?? '')?.toLocal() ??
            DateTime.now().add(const Duration(days: 30));
    final quantityTotal = requested.values.fold<double>(0, (a, b) => a + b);
    if (!demo) {
      try {
        await FirebaseFirestore.instance.runTransaction((transaction) async {
          alreadyRecorded = false;
          final recordReference = _collection(collectionName).doc(id);
          final previousRecord = await transaction.get(recordReference);
          if (previousRecord.exists) {
            alreadyRecorded = true;
            return;
          }
          final snapshots = <String, DocumentSnapshot<Map<String, dynamic>>>{};
          for (final productId in requested.keys) {
            final reference = _collection('products').doc(productId);
            snapshots[productId] = await transaction.get(reference);
          }
          total = 0;
          lineItems.clear();
          for (final item in supplied) {
            final productId = item['productId'].toString();
            final current = snapshots[productId]!;
            if (!current.exists) throw Exception('Un producto ya no existe.');
            final productData = current.data()!;
            final unit = _number(productData[purchase ? 'cost' : 'price']);
            final qty = _number(item['quantity']);
            total += unit * qty;
            lineItems.add({
              'productId': productId,
              'name': productData['name']?.toString() ?? 'Producto',
              'quantity': qty,
              'unitPrice': unit,
              'total': unit * qty,
            });
          }
          subtotal = total;
          pocketSurcharge = pocketPaymentConfirmed
              ? double.parse((subtotal * .05).toStringAsFixed(2))
              : 0;
          total = subtotal + pocketSurcharge;
          if (pocketPaymentConfirmed &&
              (_number(data['pocketAmount']) - total).abs() > .01) {
            throw Exception(
                'El total confirmado por Poket cambió. No se registró la venta; revisa el importe del enlace.');
          }
          for (final entry in requested.entries) {
            final current = snapshots[entry.key]!;
            final stock = _number(current.data()?['stock']);
            if (!purchase && stock < entry.value) {
              throw Exception(
                  'No hay suficientes existencias de ${current.data()?['name'] ?? 'un producto'}.');
            }
          }
          for (final entry in requested.entries) {
            final reference = _collection('products').doc(entry.key);
            final stock = _number(snapshots[entry.key]!.data()?['stock']);
            transaction.update(reference,
                {'stock': stock + (purchase ? entry.value : -entry.value)});
          }
          final paid = purchase
              ? total
              : data['isCredit'] == true
                  ? _number(data['amountPaid']).clamp(0, total).toDouble()
                  : data['paymentMethod'] == 'pocket'
                      ? total
                      : total;
          paymentSchedule = data['isCredit'] == true
              ? _buildPaymentSchedule(
                  balance: total - paid,
                  count: installmentCount,
                  frequency: frequency,
                  firstDueDate: firstDueDate)
              : [];
          transaction.set(recordReference, {
            ...data,
            'items': lineItems,
            'quantity': quantityTotal,
            'subtotal': subtotal,
            'pocketSurcharge': pocketSurcharge,
            'total': total,
            'amountPaid': paid,
            'balance': total - paid,
            'currency': transactionCurrency,
            'paymentSchedule': paymentSchedule,
            'paymentStatus': total - paid > 0 ? 'Pendiente' : 'Pagada',
            'date': DateFormat('yyyy-MM-dd HH:mm').format(DateTime.now())
          });
        });
      } on FirebaseException catch (error) {
        throw Exception(
            _firestoreWriteMessage(error, purchase ? 'la compra' : 'la venta'));
      }
      if (!alreadyRecorded) {
        for (final entry in requested.entries) {
          final product = products.firstWhere((p) => p['id'] == entry.key);
          product['stock'] = _number(product['stock']) +
              (purchase ? entry.value : -entry.value);
        }
      }
      final snapshot = await _collection(collectionName)
          .orderBy('date', descending: true)
          .limit(60)
          .get();
      _setList(collectionName,
          snapshot.docs.map((d) => {'id': d.id, ...d.data()}).toList());
    } else {
      total = 0;
      lineItems.clear();
      for (final item in supplied) {
        final product =
            products.firstWhere((p) => p['id'] == item['productId']);
        final qty = _number(item['quantity']);
        if (!purchase &&
            _number(product['stock']) < (requested[item['productId']] ?? 0)) {
          throw Exception(
              'No hay suficientes existencias de ${product['name']}.');
        }
        final unit = _number(product[purchase ? 'cost' : 'price']);
        total += unit * qty;
        lineItems.add({
          'productId': item['productId'],
          'name': product['name'],
          'quantity': qty,
          'unitPrice': unit,
          'total': unit * qty,
        });
      }
      subtotal = total;
      pocketSurcharge = pocketPaymentConfirmed
          ? double.parse((subtotal * .05).toStringAsFixed(2))
          : 0;
      total = subtotal + pocketSurcharge;
      if (pocketPaymentConfirmed &&
          (_number(data['pocketAmount']) - total).abs() > .01) {
        throw Exception(
            'El total confirmado por Poket cambió. No se registró la venta; revisa el importe del enlace.');
      }
      alreadyRecorded =
          _list(collectionName).any((record) => record['id']?.toString() == id);
      for (final entry in requested.entries) {
        if (alreadyRecorded) break;
        final product = products.firstWhere((p) => p['id'] == entry.key);
        product['stock'] =
            _number(product['stock']) + (purchase ? entry.value : -entry.value);
      }
      final paid = purchase
          ? total
          : data['isCredit'] == true
              ? _number(data['amountPaid']).clamp(0, total).toDouble()
              : data['paymentMethod'] == 'pocket'
                  ? total
                  : total;
      paymentSchedule = data['isCredit'] == true
          ? _buildPaymentSchedule(
              balance: total - paid,
              count: installmentCount,
              frequency: frequency,
              firstDueDate: firstDueDate)
          : [];
      if (!alreadyRecorded) {
        _list(collectionName).insert(0, {
          'id': id,
          ...data,
          'items': lineItems,
          'quantity': quantityTotal,
          'subtotal': subtotal,
          'pocketSurcharge': pocketSurcharge,
          'total': total,
          'amountPaid': paid,
          'balance': total - paid,
          'currency': transactionCurrency,
          'paymentSchedule': paymentSchedule,
          'paymentStatus': total - paid > 0 ? 'Pendiente' : 'Pagada',
          'date': DateFormat('dd/MM/yyyy HH:mm').format(DateTime.now())
        });
        await _persist('products');
        await _persist(collectionName);
      }
    }
    notifyListeners();
    if (!purchase &&
        data['isCredit'] == true &&
        data['autoReminders'] == true &&
        paymentSchedule.isNotEmpty) {
      return _syncPaymentReminders(
        saleId: id,
        saleNumber: data['number']?.toString() ?? id,
        customerName: data['party']?.toString() ?? 'cliente',
        phone: data['reminderPhone']?.toString() ?? '',
        schedule: paymentSchedule,
        enabled: true,
        currency: transactionCurrency,
      );
    }
    return true;
  }

  Future<void> attachPoketPaylink(
      String saleId, Map<String, dynamic> paylink) async {
    final update = {
      'poketPaylinkId': paylink['id']?.toString() ?? '',
      'poketPaymentLink': paylink['link']?.toString() ?? '',
      'poketLinkExpiresAt': paylink['expirationDate']?.toString() ?? '',
      'paymentStatus': 'Pendiente',
    };
    if (!demo) {
      try {
        await _collection('sales').doc(saleId).update(update);
      } catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'el enlace de pago'));
      }
    }
    final index = sales.indexWhere((sale) => sale['id']?.toString() == saleId);
    if (index >= 0) sales[index].addAll(update);
    await _persist('sales');
    notifyListeners();
  }

  Future<bool> _syncPaymentReminders({
    required String saleId,
    required String saleNumber,
    required String customerName,
    required String phone,
    required List<Map<String, dynamic>> schedule,
    required bool enabled,
    String? currency,
  }) async {
    if (demo) return false;
    const configuredApi = String.fromEnvironment('BOT_API_URL');
    final root = _resolveApiRoot(configuredApi);
    if (root.isEmpty) return false;
    try {
      final response = await http
          .put(
            Uri.parse('$root/api/reminders/plan'),
            headers: await _authorizedApiHeaders(json: true),
            body: jsonEncode({
              'saleId': saleId,
              'saleNumber': saleNumber,
              'customerName': customerName,
              'phone': phone,
              'currency': currency ?? currencyCode,
              'autoReminders': enabled,
              'paymentSchedule': schedule,
            }),
          )
          .timeout(const Duration(seconds: 20));
      return response.statusCode >= 200 && response.statusCode < 300;
    } catch (_) {
      return false;
    }
  }

  Future<bool> recordPayment(Map<String, dynamic> record, double amount) async {
    final id = record['id']?.toString();
    if (id == null || id.isEmpty) throw Exception('No encontramos la venta.');
    final remaining = _number(record['balance']);
    if (amount <= 0 || amount > remaining) {
      throw Exception(
          'El abono debe ser mayor que cero y no superar el saldo.');
    }
    final now = DateTime.now();
    final payment = {
      'amount': amount,
      'date': DateFormat('yyyy-MM-dd HH:mm').format(now),
    };
    var paid = _number(record['amountPaid']) + amount;
    var balance = remaining - amount;
    var updatedSchedule =
        _applyInstallmentPayment(record['paymentSchedule'], amount);
    if (!demo) {
      final reference = _collection('sales').doc(id);
      try {
        await FirebaseFirestore.instance.runTransaction((transaction) async {
          final snapshot = await transaction.get(reference);
          if (!snapshot.exists) throw Exception('La venta ya no existe.');
          final current = snapshot.data()!;
          final currentBalance = _number(current['balance']);
          if (amount > currentBalance) {
            throw Exception(
                'El saldo cambió. Actualiza las ventas e inténtalo de nuevo.');
          }
          final nextPaid = _number(current['amountPaid']) + amount;
          final nextBalance = currentBalance - amount;
          paid = nextPaid;
          balance = nextBalance;
          updatedSchedule =
              _applyInstallmentPayment(current['paymentSchedule'], amount);
          transaction.update(reference, {
            'amountPaid': nextPaid,
            'balance': nextBalance,
            'paymentStatus': nextBalance <= 0 ? 'Pagada' : 'Pendiente',
            'payments': FieldValue.arrayUnion([payment]),
            'paymentSchedule': updatedSchedule,
          });
        });
      } on FirebaseException catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'el abono'));
      }
    }
    final index = sales.indexWhere((sale) => sale['id']?.toString() == id);
    if (index >= 0) {
      final updated = Map<String, dynamic>.from(sales[index]);
      updated['amountPaid'] = paid;
      updated['balance'] = balance;
      updated['paymentStatus'] = balance <= 0 ? 'Pagada' : 'Pendiente';
      updated['payments'] = [...(updated['payments'] as List? ?? []), payment];
      updated['paymentSchedule'] = updatedSchedule;
      sales[index] = updated;
    }
    if (demo) await _persist('sales');
    notifyListeners();
    if (record['autoReminders'] == true) {
      return _syncPaymentReminders(
        saleId: id,
        saleNumber: record['number']?.toString() ?? id,
        customerName: record['party']?.toString() ?? 'cliente',
        phone: record['reminderPhone']?.toString() ?? '',
        schedule: updatedSchedule,
        enabled: true,
        currency: record['currency']?.toString(),
      );
    }
    return true;
  }

  Future<Map<String, dynamic>> recordSaleReturn({
    required String saleId,
    required List<Map<String, dynamic>> items,
    required String reason,
    required String refundMethod,
  }) async {
    if (saleId.isEmpty) throw Exception('No encontramos la venta.');
    if (items.isEmpty) throw Exception('Selecciona productos para devolver.');
    final returnId = DateTime.now().microsecondsSinceEpoch.toString();
    final now = DateTime.now();
    Map<String, dynamic> resultRecord = {};
    var syncReminders = false;
    if (demo) {
      final saleIndex =
          sales.indexWhere((sale) => sale['id']?.toString() == saleId);
      if (saleIndex < 0) throw Exception('No encontramos la venta.');
      final groupedIds =
          items.map((item) => item['productId']?.toString() ?? '').toSet();
      final productsById = <String, Map<String, dynamic>>{
        for (final product in products
            .where((product) => groupedIds.contains(product['id']?.toString())))
          product['id'].toString(): product,
      };
      syncReminders = sales[saleIndex]['autoReminders'] == true;
      final calculated = _calculateSaleReturn(
        sale: sales[saleIndex],
        requestedItems: items,
        productsById: productsById,
        reason: reason,
        refundMethod: refundMethod,
        returnId: returnId,
        at: now,
      );
      final additions =
          Map<String, double>.from(calculated['stockAdditions'] as Map);
      for (final entry in additions.entries) {
        final product = productsById[entry.key]!;
        product['stock'] = _number(product['stock']) + entry.value;
      }
      sales[saleIndex]
          .addAll(Map<String, dynamic>.from(calculated['saleUpdate'] as Map));
      resultRecord =
          Map<String, dynamic>.from(calculated['returnRecord'] as Map);
      returns = [resultRecord, ...returns];
      await _persist('products');
      await _persist('sales');
      await _persist('returns');
    } else {
      final saleReference = _collection('sales').doc(saleId);
      final returnReference = _collection('returns').doc(returnId);
      try {
        await FirebaseFirestore.instance.runTransaction((transaction) async {
          final saleSnapshot = await transaction.get(saleReference);
          if (!saleSnapshot.exists) throw Exception('No encontramos la venta.');
          final sale = {'id': saleSnapshot.id, ...saleSnapshot.data()!};
          syncReminders = sale['autoReminders'] == true;
          final productIds = items
              .map((item) => item['productId']?.toString() ?? '')
              .where((id) => id.isNotEmpty)
              .toSet();
          final productSnapshots =
              <String, DocumentSnapshot<Map<String, dynamic>>>{};
          for (final id in productIds) {
            productSnapshots[id] =
                await transaction.get(_collection('products').doc(id));
          }
          final productsById = <String, Map<String, dynamic>>{
            for (final entry in productSnapshots.entries)
              if (entry.value.exists)
                entry.key: {'id': entry.key, ...entry.value.data()!},
          };
          final calculated = _calculateSaleReturn(
            sale: sale,
            requestedItems: items,
            productsById: productsById,
            reason: reason,
            refundMethod: refundMethod,
            returnId: returnId,
            at: now,
          );
          final additions =
              Map<String, double>.from(calculated['stockAdditions'] as Map);
          for (final entry in additions.entries) {
            final snapshot = productSnapshots[entry.key]!;
            final stock = _number(snapshot.data()?['stock']);
            transaction.update(_collection('products').doc(entry.key),
                {'stock': stock + entry.value});
          }
          transaction.update(
              saleReference, calculated['saleUpdate'] as Map<String, dynamic>);
          final returnRecord =
              Map<String, dynamic>.from(calculated['returnRecord'] as Map);
          transaction.set(returnReference, returnRecord);
          resultRecord = returnRecord;
        });
        final saleSnapshot = await _collection('sales')
            .orderBy('date', descending: true)
            .limit(100)
            .get();
        _setList(
            'sales',
            saleSnapshot.docs
                .map((doc) => {'id': doc.id, ...doc.data()})
                .toList());
        final productSnapshot = await _collection('products').get();
        _setList(
            'products',
            productSnapshot.docs
                .map((doc) => {'id': doc.id, ...doc.data()})
                .toList());
        final returnSnapshot = await _collection('returns')
            .orderBy('date', descending: true)
            .limit(100)
            .get();
        _setList(
            'returns',
            returnSnapshot.docs
                .map((doc) => {'id': doc.id, ...doc.data()})
                .toList());
      } on FirebaseException catch (error) {
        throw Exception(_firestoreWriteMessage(error, 'la devolución'));
      }
    }
    final updatedSale =
        sales.where((sale) => sale['id']?.toString() == saleId).firstOrNull;
    if (!demo && syncReminders) {
      final schedule = (updatedSale?['paymentSchedule'] as List? ?? [])
          .whereType<Map>()
          .map((item) => Map<String, dynamic>.from(item))
          .toList();
      await _syncPaymentReminders(
        saleId: saleId,
        saleNumber: updatedSale?['number']?.toString() ?? saleId,
        customerName: updatedSale?['party']?.toString() ?? 'cliente',
        phone: updatedSale?['reminderPhone']?.toString() ?? '',
        schedule: schedule,
        enabled: updatedSale?['autoReminders'] == true,
        currency: updatedSale?['currency']?.toString(),
      );
    }
    notifyListeners();
    return resultRecord;
  }
}

class _PoketPage extends StatefulWidget {
  const _PoketPage();
  @override
  State<_PoketPage> createState() => _PoketPageState();
}

class _PoketPageState extends State<_PoketPage> {
  static const _configuredUrl = String.fromEnvironment('BOT_API_URL');
  String get _root => _resolveApiRoot(_configuredUrl);
  final _apiKey = TextEditingController();
  final _merchantId = TextEditingController();
  final _terminalId = TextEditingController();
  final _paymentUrl = TextEditingController();
  bool _loading = true, _saving = false, _checking = false, _deleting = false;
  bool _configured = false, _hasSavedKey = false;
  String _message = '', _error = '';

  @override
  void initState() {
    super.initState();
    _load();
  }

  @override
  void dispose() {
    _apiKey.dispose();
    _merchantId.dispose();
    _terminalId.dispose();
    _paymentUrl.dispose();
    super.dispose();
  }

  Future<void> _load() async {
    if (_root.isEmpty) {
      if (mounted) {
        setState(() {
          _loading = false;
          _error = 'Inicia el servidor de Sync Connect para vincular Poket.';
        });
      }
      return;
    }
    try {
      final response = await http
          .get(Uri.parse('$_root/api/state'),
              headers: await _authorizedApiHeaders())
          .timeout(const Duration(seconds: 12));
      final data = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300) {
        throw Exception(data['error'] ?? 'No se pudo leer el estado de Poket.');
      }
      final business =
          (data['business'] as Map?)?.cast<String, dynamic>() ?? {};
      _merchantId.text = data['poketMerchantId']?.toString() ?? '';
      _terminalId.text = data['poketTerminalId']?.toString() ?? '';
      _paymentUrl.text = business['poketPaymentUrl']?.toString() ?? '';
      if (mounted) {
        setState(() {
          _hasSavedKey = data['poketApiKeyConfigured'] == true;
          _configured = _hasSavedKey &&
              _merchantId.text.isNotEmpty &&
              _terminalId.text.isNotEmpty;
          final credentialError =
              data['poketCredentialError']?.toString() ?? '';
          if (credentialError.isNotEmpty) _error = credentialError;
          _loading = false;
        });
      }
    } catch (error) {
      if (mounted) {
        setState(() {
          _loading = false;
          _error = '$error'.replaceFirst('Exception: ', '');
        });
      }
    }
  }

  Future<void> _save() async {
    if (_saving || _deleting) return;
    final key = _apiKey.text.trim();
    final merchant = _merchantId.text.trim();
    final terminal = _terminalId.text.trim();
    if (!_hasSavedKey && key.isEmpty) {
      setState(() =>
          _error = 'Coloca tu API key (PAT) de Poket para vincular la cuenta.');
      return;
    }
    if (merchant.isEmpty || terminal.isEmpty) {
      setState(() => _error =
          'Poket requiere Merchant ID y Terminal ID de tipo Ecommerce.');
      return;
    }
    final apiRoot = Uri.tryParse(_root);
    final isLocalApi = apiRoot != null &&
        const {'localhost', '127.0.0.1', '::1'}.contains(apiRoot.host);
    if (apiRoot == null || (apiRoot.scheme != 'https' && !isLocalApi)) {
      setState(() => _error =
          'Por seguridad, usa HTTPS o el servidor local para guardar la API key.');
      return;
    }
    setState(() {
      _saving = true;
      _error = '';
      _message = '';
    });
    try {
      final settingsResponse = await http
          .put(Uri.parse('$_root/api/settings'),
              headers: await _authorizedApiHeaders(json: true),
              body: jsonEncode({
                'business': {'poketPaymentUrl': _paymentUrl.text.trim()}
              }))
          .timeout(const Duration(seconds: 15));
      final settingsBody =
          jsonDecode(settingsResponse.body) as Map<String, dynamic>;
      if (settingsResponse.statusCode >= 300) {
        throw Exception(
            settingsBody['error'] ?? 'No se guardó la configuración.');
      }
      final response = await http
          .put(Uri.parse('$_root/api/poket/credentials'),
              headers: await _authorizedApiHeaders(json: true),
              body: jsonEncode({
                'apiKey': key,
                'merchantId': merchant,
                'terminalId': terminal,
              }))
          .timeout(const Duration(seconds: 15));
      final data = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300) {
        throw Exception(data['error'] ?? 'Poket no pudo guardar la conexión.');
      }
      _apiKey.clear();
      if (mounted) {
        setState(() {
          _configured = data['configured'] == true;
          _hasSavedKey = data['configured'] == true;
          _message =
              'Conexión guardada y API key cifrada. Pulsa “Probar conexión” para verificarla con Poket.';
        });
        ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(content: Text('Configuración de Poket guardada.')));
      }
    } catch (error) {
      if (mounted) {
        setState(() => _error = '$error'.replaceFirst('Exception: ', ''));
      }
    } finally {
      if (mounted) setState(() => _saving = false);
    }
  }

  Future<void> _checkConnection() async {
    if (!_configured || _checking || _saving) return;
    setState(() {
      _checking = true;
      _error = '';
      _message = 'Verificando credenciales con Poket…';
    });
    try {
      final response = await http
          .post(Uri.parse('$_root/api/poket/validate'),
              headers: await _authorizedApiHeaders(json: true), body: '{}')
          .timeout(const Duration(seconds: 20));
      final data = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300) {
        throw Exception(data['error'] ?? 'No se pudo validar la conexión.');
      }
      if (mounted)
        setState(() => _message =
            data['message']?.toString() ?? 'Poket respondió correctamente.');
    } catch (error) {
      if (mounted) {
        setState(() {
          _message = '';
          _error = '$error'.replaceFirst('Exception: ', '');
        });
      }
    } finally {
      if (mounted) setState(() => _checking = false);
    }
  }

  Future<void> _delete() async {
    if (!_hasSavedKey || _deleting || _saving) return;
    setState(() {
      _deleting = true;
      _error = '';
    });
    try {
      final response = await http
          .delete(Uri.parse('$_root/api/poket/credentials'),
              headers: await _authorizedApiHeaders())
          .timeout(const Duration(seconds: 15));
      final data = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300) {
        throw Exception(data['error'] ?? 'No se pudo desvincular Poket.');
      }
      _apiKey.clear();
      _merchantId.clear();
      _terminalId.clear();
      if (mounted) {
        setState(() {
          _configured = false;
          _hasSavedKey = false;
          _message = '';
        });
        ScaffoldMessenger.of(context).showSnackBar(const SnackBar(
            content: Text('Poket se desvinculó de esta cuenta.')));
      }
    } catch (error) {
      if (mounted)
        setState(() => _error = '$error'.replaceFirst('Exception: ', ''));
    } finally {
      if (mounted) setState(() => _deleting = false);
    }
  }

  Future<void> _openPortal() async {
    final opened = await launchUrl(
        Uri.parse('https://portal.pagoconpoket.com/'),
        mode: LaunchMode.externalApplication);
    if (!opened && mounted) {
      ScaffoldMessenger.of(context).showSnackBar(const SnackBar(
          content: Text('No se pudo abrir el portal de Poket.')));
    }
  }

  Widget _field(TextEditingController controller, String label,
          {String? helper, bool secret = false}) =>
      Padding(
        padding: const EdgeInsets.only(bottom: 12),
        child: TextField(
          controller: controller,
          obscureText: secret,
          autocorrect: !secret,
          enableSuggestions: !secret,
          decoration: InputDecoration(labelText: label, helperText: helper),
        ),
      );

  @override
  Widget build(BuildContext context) {
    if (_loading) return const _Loading();
    return ListView(
      padding: const EdgeInsets.all(24),
      children: [
        const _PageHeading('Vincular Poket',
            'Conecta tu cuenta para crear enlaces de pago individuales y enviarlos por WhatsApp.'),
        const SizedBox(height: 16),
        Align(
          alignment: Alignment.topCenter,
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 760),
            child: Card(
              elevation: 0,
              child: Padding(
                padding: const EdgeInsets.all(22),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(children: [
                      Container(
                        width: 46,
                        height: 46,
                        decoration: BoxDecoration(
                            color: const Color(0xffedf3ff),
                            borderRadius: BorderRadius.circular(14)),
                        child: const Icon(Icons.account_balance_wallet_outlined,
                            color: _brand),
                      ),
                      const SizedBox(width: 12),
                      const Expanded(
                          child: Text('Cuenta de Poket',
                              style: TextStyle(
                                  fontSize: 19, fontWeight: FontWeight.w800))),
                      if (_configured)
                        const Chip(
                          avatar: Icon(Icons.check_circle_outline, size: 17),
                          label: Text('Vinculada'),
                        ),
                    ]),
                    const SizedBox(height: 14),
                    const Text(
                        'Ingresa la API key (PAT), el Merchant ID y la terminal Ecommerce de tu comercio. Encontrarás los IDs en el portal de Poket, en tu perfil → Datos de mi comercio.'),
                    const SizedBox(height: 18),
                    _field(
                      _apiKey,
                      'API key (PAT)',
                      secret: true,
                      helper: _hasSavedKey
                          ? 'La clave ya está cifrada. Déjala vacía para conservarla o escribe una nueva.'
                          : 'Se cifra en el servidor y no vuelve a mostrarse.',
                    ),
                    _field(_merchantId, 'Merchant ID'),
                    _field(_terminalId, 'Terminal ID (Ecommerce)'),
                    _field(
                      _paymentUrl,
                      'Enlace fijo de cobro (opcional)',
                      helper:
                          'Se usa como alternativa manual. Los cobros por venta crean enlaces por monto.',
                    ),
                    Align(
                      alignment: Alignment.centerLeft,
                      child: TextButton.icon(
                        onPressed: _openPortal,
                        icon: const Icon(Icons.open_in_new),
                        label: const Text('Abrir portal de Poket'),
                      ),
                    ),
                    if (_message.isNotEmpty)
                      Padding(
                        padding: const EdgeInsets.only(top: 8),
                        child: Text(_message,
                            style: const TextStyle(color: Color(0xff187348))),
                      ),
                    if (_error.isNotEmpty)
                      Padding(
                        padding: const EdgeInsets.only(top: 8),
                        child: SelectableText(_error,
                            style: const TextStyle(color: Colors.red)),
                      ),
                    const SizedBox(height: 16),
                    Wrap(spacing: 10, runSpacing: 8, children: [
                      FilledButton.icon(
                        onPressed:
                            _saving || _checking || _deleting ? null : _save,
                        icon: Icon(
                            _saving ? Icons.hourglass_top : Icons.link_rounded),
                        label: Text(_saving
                            ? 'Guardando…'
                            : _configured
                                ? 'Guardar configuración'
                                : 'Vincular aplicación de Poket'),
                      ),
                      if (_configured)
                        OutlinedButton.icon(
                          onPressed: _checking || _saving || _deleting
                              ? null
                              : _checkConnection,
                          icon: Icon(_checking
                              ? Icons.hourglass_top
                              : Icons.wifi_tethering_rounded),
                          label: Text(
                              _checking ? 'Comprobando…' : 'Probar conexión'),
                        ),
                      if (_hasSavedKey)
                        TextButton.icon(
                          onPressed: _deleting || _saving ? null : _delete,
                          icon: Icon(
                              _deleting ? Icons.hourglass_top : Icons.link_off),
                          label: Text(
                              _deleting ? 'Desvinculando…' : 'Desvincular'),
                        ),
                    ]),
                    const SizedBox(height: 14),
                    const Text(
                        'Los PayLinks son de un solo uso y vencen en siete días. Las monedas habilitadas en esta integración son NIO, HNL y CRC.',
                        style: TextStyle(color: Colors.black54)),
                  ],
                ),
              ),
            ),
          ),
        ),
      ],
    );
  }
}

class _ConnectionsPage extends StatefulWidget {
  const _ConnectionsPage({required this.store, required this.onNavigate});
  final _Store store;
  final ValueChanged<int> onNavigate;
  @override
  State<_ConnectionsPage> createState() => _ConnectionsPageState();
}

class _ConnectionsPageState extends State<_ConnectionsPage> {
  static const _configuredUrl = String.fromEnvironment('BOT_API_URL');
  String get _root => _resolveApiRoot(_configuredUrl);
  final _name = TextEditingController();
  final _description = TextEditingController();
  final _goal = TextEditingController();
  final _tone = TextEditingController();
  final _services = TextEditingController();
  final _hours = TextEditingController();
  final _address = TextEditingController();
  final _prices = TextEditingController();
  final _faqs = TextEditingController();
  final _quote = TextEditingController();
  final _phone = TextEditingController();
  final _taxId = TextEditingController();
  final _receiptAddress = TextEditingController();
  final _receiptPhone = TextEditingController();
  final _receiptEmail = TextEditingController();
  bool _loading = true, _saving = false, _exporting = false;
  bool _savingReceipt = false;
  bool _isAdmin = false;
  int _globalAiCount = 0;
  String _aiSummary = 'La IA se administra centralmente.';
  String _waStatus = 'Consultando conexión…', _waError = '';
  late String _selectedCurrency = widget.store.currencyCode;
  @override
  void initState() {
    super.initState();
    _load();
  }

  @override
  void dispose() {
    for (final c in [
      _name,
      _description,
      _goal,
      _tone,
      _services,
      _hours,
      _address,
      _prices,
      _faqs,
      _quote,
      _phone,
      _taxId,
      _receiptAddress,
      _receiptPhone,
      _receiptEmail,
    ]) {
      c.dispose();
    }
    super.dispose();
  }

  Future<void> _load() async {
    if (_root.isEmpty) {
      if (mounted)
        setState(() {
          _loading = false;
          _waError = 'Inicia Sync Connect en este computador.';
        });
      return;
    }
    try {
      final r = await http
          .get(Uri.parse('$_root/api/state'),
              headers: await _authorizedApiHeaders())
          .timeout(const Duration(seconds: 8));
      final d = jsonDecode(r.body) as Map<String, dynamic>;
      final b = (d['business'] as Map?)?.cast<String, dynamic>() ?? {};
      _name.text = b['name']?.toString() ?? '';
      _description.text = b['description']?.toString() ?? '';
      _goal.text = b['goal']?.toString() ?? '';
      _tone.text = b['compatibility']?.toString() ?? '';
      _services.text = b['services']?.toString() ?? '';
      _hours.text = b['hours']?.toString() ?? '';
      _address.text = b['address']?.toString() ?? '';
      _prices.text = b['prices']?.toString() ?? '';
      _faqs.text = b['faqs']?.toString() ?? '';
      _quote.text = b['quoteMessage']?.toString() ?? '';
      _phone.text = b['notificationPhone']?.toString() ?? '';
      final receipt = widget.store.receiptConfig;
      _taxId.text = receipt['taxId']?.toString() ?? '';
      _receiptAddress.text = receipt['address']?.toString() ?? '';
      _receiptPhone.text = receipt['phone']?.toString() ?? '';
      _receiptEmail.text = receipt['email']?.toString() ?? '';
      _selectedCurrency = widget.store.currencyCode;
      if (mounted)
        setState(() {
          _isAdmin = d['isAdmin'] == true;
          _globalAiCount = _number(d['globalAiProviderCount']).toInt();
          _aiSummary = _globalAiCount == 0
              ? (d['apiConfigured'] == true
                  ? 'Hay un proveedor de IA disponible.'
                  : 'El administrador todavía no ha activado proveedores de IA.')
              : '$_globalAiCount proveedor(es) configurado(s) por administración.';
          _waStatus = d['waStatus']?.toString() ?? 'Sin estado';
          _waError = d['waLastError']?.toString() ?? '';
          _loading = false;
        });
    } catch (e) {
      if (mounted)
        setState(() {
          _loading = false;
          _waError = 'No se pudo leer la configuración: $e';
        });
    }
  }

  Future<void> _save() async {
    if (_saving) return;
    setState(() {
      _saving = true;
      _waError = '';
    });
    try {
      final r = await http
          .put(Uri.parse('$_root/api/settings'),
              headers: await _authorizedApiHeaders(json: true),
              body: jsonEncode({
                'business': {
                  'name': _name.text.trim(),
                  'description': _description.text.trim(),
                  'goal': _goal.text.trim(),
                  'compatibility': _tone.text.trim(),
                  'services': _services.text.trim(),
                  'hours': _hours.text.trim(),
                  'address': _address.text.trim(),
                  'prices': _prices.text.trim(),
                  'faqs': _faqs.text.trim(),
                  'quoteMessage': _quote.text.trim(),
                  'notificationPhone': _phone.text.trim(),
                  'currency': _selectedCurrency,
                },
              }))
          .timeout(const Duration(seconds: 15));
      final d = jsonDecode(r.body) as Map<String, dynamic>;
      if (r.statusCode >= 300)
        throw Exception(d['error'] ?? 'No se pudo guardar.');
      await _load();
      if (mounted)
        ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(content: const Text('Configuración guardada.')));
    } catch (e) {
      if (mounted) setState(() => _waError = '$e');
    } finally {
      if (mounted) setState(() => _saving = false);
    }
  }

  Future<void> _saveReceipt() async {
    if (_savingReceipt) return;
    setState(() => _savingReceipt = true);
    try {
      await widget.store.updateReceiptConfig({
        'businessName': _name.text.trim(),
        'taxId': _taxId.text.trim(),
        'address': _receiptAddress.text.trim(),
        'phone': _receiptPhone.text.trim(),
        'email': _receiptEmail.text.trim(),
      });
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Datos del comprobante guardados.')),
        );
      }
    } catch (error) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('$error'.replaceFirst('Exception: ', ''))),
        );
      }
    } finally {
      if (mounted) setState(() => _savingReceipt = false);
    }
  }

  xls.CellValue? _excelCell(dynamic value) {
    if (value == null) return null;
    if (value is bool) return xls.BoolCellValue(value);
    if (value is int) return xls.IntCellValue(value);
    if (value is num) return xls.DoubleCellValue(value.toDouble());
    if (value is Timestamp) {
      return xls.TextCellValue(value.toDate().toIso8601String());
    }
    if (value is DateTime) return xls.TextCellValue(value.toIso8601String());
    final text = value is Map || value is Iterable
        ? jsonEncode(value, toEncodable: (item) {
            if (item is Timestamp) return item.toDate().toIso8601String();
            if (item is DateTime) return item.toIso8601String();
            return item.toString();
          })
        : value.toString();
    return xls.TextCellValue(text.length > 32760
        ? '${text.substring(0, 32740)} [texto largo]'
        : text);
  }

  List<Map<String, dynamic>> _exportMaps(dynamic value) =>
      (value is List ? value : const [])
          .whereType<Map>()
          .map((item) => Map<String, dynamic>.from(item))
          .toList();

  List<String> _exportColumns(List<Map<String, dynamic>> rows,
      {Set<String> omit = const {}}) {
    final columns = <String>{};
    for (final row in rows) {
      columns.addAll(row.keys.where((key) => !omit.contains(key)));
    }
    return columns.toList();
  }

  void _writeExcelSheet(xls.Excel workbook, String name, List<String> headers,
      Iterable<List<dynamic>> rows) {
    if (headers.isEmpty) headers = const ['Sin registros'];
    final sheet = workbook[name];
    final headerStyle = xls.CellStyle(
      backgroundColorHex: xls.ExcelColor.blue,
      fontColorHex: xls.ExcelColor.white,
      bold: true,
    );
    sheet.appendRow(headers.map((value) => xls.TextCellValue(value)).toList());
    for (var i = 0; i < headers.length; i++) {
      sheet.updateCell(
        xls.CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0),
        xls.TextCellValue(headers[i]),
        cellStyle: headerStyle,
      );
    }
    for (final row in rows) {
      sheet.appendRow(row.map(_excelCell).toList());
    }
    if (headers.isNotEmpty) sheet.setColumnAutoFit(0);
  }

  String _exportImageValue(List<List<dynamic>> imageRows, String type,
      String id, String name, dynamic raw) {
    final value = raw?.toString() ?? '';
    final match = RegExp(r'^data:(image/[^;]+);base64,', caseSensitive: false)
        .firstMatch(value);
    if (match == null) return value;
    final encoded = value.substring(match.end);
    const chunkLength = 30000;
    final total = (encoded.length / chunkLength).ceil();
    for (var offset = 0; offset < encoded.length; offset += chunkLength) {
      final part = (offset ~/ chunkLength) + 1;
      imageRows.add([
        type,
        id,
        name,
        match.group(1) ?? 'image/unknown',
        part,
        total,
        encoded.substring(
            offset, math.min(offset + chunkLength, encoded.length)),
      ]);
    }
    return 'Imagen en hoja Imágenes: $type / $id';
  }

  Future<List<Map<String, dynamic>>> _loadExportCollection(String name) async {
    if (widget.store.demo) return widget.store._list(name);
    final snapshot = await FirebaseFirestore.instance
        .collection('users')
        .doc(widget.store.uid)
        .collection(name)
        .get()
        .timeout(const Duration(seconds: 45));
    return snapshot.docs.map((doc) => {'id': doc.id, ...doc.data()}).toList();
  }

  Future<void> _exportWorkbook() async {
    if (_exporting) return;
    setState(() => _exporting = true);
    try {
      final names = const [
        'products',
        'contacts',
        'sales',
        'returns',
        'purchases',
        'expenses'
      ];
      final loaded = await Future.wait(names.map(_loadExportCollection));
      final records = <String, List<Map<String, dynamic>>>{
        for (var i = 0; i < names.length; i++) names[i]: loaded[i],
      };

      var remote = <String, dynamic>{};
      var remoteStatus = widget.store.demo
          ? 'Modo local: no hay conversaciones de WhatsApp sincronizadas.'
          : 'Servicio de WhatsApp no disponible; se exportaron los datos de Firebase.';
      if (!widget.store.demo && _root.isNotEmpty) {
        try {
          final response = await http
              .get(Uri.parse('$_root/api/export'),
                  headers: await _authorizedApiHeaders())
              .timeout(const Duration(seconds: 25));
          final body = jsonDecode(response.body);
          if (response.statusCode >= 300 || body is! Map) {
            final message = body is Map ? body['error'] : null;
            throw Exception(
                message ?? 'No se pudo leer el servicio de WhatsApp.');
          }
          remote = Map<String, dynamic>.from(body);
          remoteStatus =
              'Conversaciones, contactos, catálogos y avisos incluidos.';
        } catch (error) {
          remoteStatus = 'WhatsApp y catálogos no disponibles: $error';
        }
      }

      final workbook = xls.Excel.createExcel();
      final defaultSheet = workbook.getDefaultSheet();
      if (defaultSheet != null) workbook.rename(defaultSheet, 'Resumen');
      final imageRows = <List<dynamic>>[];
      final now = DateTime.now();
      final summaryRows = <List<dynamic>>[
        ['Fecha de exportación', now.toIso8601String()],
        ['Cuenta', FirebaseAuth.instance.currentUser?.email ?? 'Cuenta local'],
        ['Moneda', widget.store.currencyCode],
        ['Productos', records['products']!.length],
        ['Clientes', records['contacts']!.length],
        ['Ventas', records['sales']!.length],
        ['Devoluciones', records['returns']!.length],
        ['Compras', records['purchases']!.length],
        ['Gastos', records['expenses']!.length],
        ['Estado de datos de mensajería', remoteStatus],
      ];
      _writeExcelSheet(workbook, 'Resumen', ['Dato', 'Valor'], summaryRows);

      final products = records['products']!;
      final inventoryColumns = _exportColumns(products);
      if (!inventoryColumns.contains('image')) inventoryColumns.add('image');
      final productRows = products.map((item) {
        final row = Map<String, dynamic>.from(item);
        row['image'] = _exportImageValue(
            imageRows,
            'producto',
            row['id']?.toString() ?? '',
            row['name']?.toString() ?? '',
            row['image']);
        return inventoryColumns.map((key) => row[key]).toList();
      }).toList();
      _writeExcelSheet(workbook, 'Inventario', inventoryColumns, productRows);
      final contactColumns = _exportColumns(records['contacts']!);
      _writeExcelSheet(
          workbook,
          'Clientes',
          contactColumns,
          records['contacts']!
              .map((row) => contactColumns.map((key) => row[key]).toList()));

      final sales = records['sales']!;
      final salesColumns = _exportColumns(sales,
          omit: const {'items', 'payments', 'paymentSchedule'});
      _writeExcelSheet(workbook, 'Ventas', salesColumns,
          sales.map((row) => salesColumns.map((key) => row[key]).toList()));
      final saleDetails = <List<dynamic>>[];
      final payments = <List<dynamic>>[];
      final installments = <List<dynamic>>[];
      for (final sale in sales) {
        final saleId = sale['id']?.toString() ?? '';
        for (final item in _exportMaps(sale['items'])) {
          saleDetails.add([
            saleId,
            sale['number'],
            sale['date'],
            sale['party'],
            item['productId'],
            item['name'],
            item['quantity'],
            item['unitPrice'],
            item['total'],
            sale['currency'],
          ]);
        }
        for (final payment in _exportMaps(sale['payments'])) {
          payments
              .add([saleId, sale['number'], sale['party'], ...payment.values]);
        }
        for (final installment in _exportMaps(sale['paymentSchedule'])) {
          installments.add(
              [saleId, sale['number'], sale['party'], ...installment.values]);
        }
      }
      _writeExcelSheet(
          workbook,
          'Detalle ventas',
          [
            'Venta ID',
            'Número',
            'Fecha',
            'Cliente',
            'Producto ID',
            'Producto',
            'Cantidad',
            'Precio unitario',
            'Total línea',
            'Moneda'
          ],
          saleDetails);
      _writeExcelSheet(workbook, 'Abonos',
          ['Venta ID', 'Número', 'Cliente', 'Importe', 'Fecha'], payments);
      _writeExcelSheet(
          workbook,
          'Cuotas',
          [
            'Venta ID',
            'Número',
            'Cliente',
            'Número de cuota',
            'Monto',
            'Pagado',
            'Pendiente',
            'Vencimiento',
            'Estado'
          ],
          installments);

      final returns = records['returns']!;
      final returnColumns = _exportColumns(returns, omit: const {'items'});
      _writeExcelSheet(
        workbook,
        'Devoluciones',
        returnColumns,
        returns.map((row) => returnColumns.map((key) => row[key]).toList()),
      );
      final returnDetails = <List<dynamic>>[];
      for (final returnRecord in returns) {
        for (final item in _exportMaps(returnRecord['items'])) {
          returnDetails.add([
            returnRecord['id'],
            returnRecord['saleNumber'],
            returnRecord['date'],
            returnRecord['customer'],
            item['productId'],
            item['name'],
            item['quantity'],
            item['unitPrice'],
            item['total'],
            returnRecord['refundAmount'],
            returnRecord['refundMethod'],
            returnRecord['reason'],
            returnRecord['currency'],
          ]);
        }
      }
      _writeExcelSheet(
        workbook,
        'Detalle devoluciones',
        [
          'Devolución ID',
          'Venta',
          'Fecha',
          'Cliente',
          'Producto ID',
          'Producto',
          'Cantidad',
          'Precio unitario',
          'Importe',
          'Reembolso',
          'Medio de reembolso',
          'Motivo',
          'Moneda',
        ],
        returnDetails,
      );

      final purchases = records['purchases']!;
      final purchaseColumns = _exportColumns(purchases, omit: const {'items'});
      _writeExcelSheet(
          workbook,
          'Compras',
          purchaseColumns,
          purchases
              .map((row) => purchaseColumns.map((key) => row[key]).toList()));
      final purchaseDetails = <List<dynamic>>[];
      for (final purchase in purchases) {
        for (final item in _exportMaps(purchase['items'])) {
          purchaseDetails.add([
            purchase['id'],
            purchase['number'],
            purchase['date'],
            purchase['party'],
            item['productId'],
            item['name'],
            item['quantity'],
            item['unitPrice'],
            item['total'],
            purchase['currency'],
          ]);
        }
      }
      _writeExcelSheet(
          workbook,
          'Detalle compras',
          [
            'Compra ID',
            'Número',
            'Fecha',
            'Proveedor',
            'Producto ID',
            'Producto',
            'Cantidad',
            'Costo unitario',
            'Total línea',
            'Moneda'
          ],
          purchaseDetails);
      _writeExcelSheet(
          workbook,
          'Gastos',
          _exportColumns(records['expenses']!),
          records['expenses']!.map((row) => _exportColumns(records['expenses']!)
              .map((key) => row[key])
              .toList()));

      final business = Map<String, dynamic>.from(remote['business'] is Map
          ? remote['business'] as Map
          : <String, dynamic>{});
      business.putIfAbsent('name', () => _name.text.trim());
      business.putIfAbsent('description', () => _description.text.trim());
      business.putIfAbsent('services', () => _services.text.trim());
      business.putIfAbsent('prices', () => _prices.text.trim());
      business.putIfAbsent('hours', () => _hours.text.trim());
      business.putIfAbsent('address', () => _address.text.trim());
      business.putIfAbsent('faqs', () => _faqs.text.trim());
      business.putIfAbsent('notificationPhone', () => _phone.text.trim());
      business['moneda'] = widget.store.currencyCode;
      final catalogConfig =
          Map<String, dynamic>.from(widget.store.catalogConfig);
      for (final entry in catalogConfig.entries) {
        final value = entry.key.toLowerCase().contains('logo') ||
                entry.key.toLowerCase().contains('image')
            ? _exportImageValue(imageRows, 'configuración catálogo', entry.key,
                entry.key, entry.value)
            : entry.value;
        business['catalog.${entry.key}'] = value;
      }
      final configRows =
          business.entries.map((entry) => [entry.key, entry.value]);
      _writeExcelSheet(
          workbook, 'Configuración', ['Campo', 'Valor'], configRows);

      final waContacts = _exportMaps(remote['whatsappContacts']);
      _writeExcelSheet(
          workbook,
          'Contactos WA',
          _exportColumns(waContacts),
          waContacts.map((row) =>
              _exportColumns(waContacts).map((key) => row[key]).toList()));
      final conversations = _exportMaps(remote['conversations']);
      final chatRows = <List<dynamic>>[];
      for (final chat in conversations) {
        final messages = _exportMaps(chat['messages']);
        for (final message in messages) {
          chatRows.add([
            chat['id'],
            chat['name'],
            message['from'],
            message['text'],
            message['at'],
            jsonEncode(message),
          ]);
        }
      }
      _writeExcelSheet(
          workbook,
          'Conversaciones WA',
          [
            'Chat ID',
            'Contacto',
            'Emisor',
            'Mensaje',
            'Fecha',
            'Datos completos'
          ],
          chatRows);
      final questions = _exportMaps(remote['unknownQuestions']);
      _writeExcelSheet(
          workbook,
          'Preguntas bot',
          _exportColumns(questions),
          questions.map((row) =>
              _exportColumns(questions).map((key) => row[key]).toList()));

      final catalogs = _exportMaps(remote['catalogs']);
      final catalogHeaders =
          _exportColumns(catalogs, omit: const {'products', 'logo'});
      _writeExcelSheet(
          workbook,
          'Catálogos',
          catalogHeaders,
          catalogs.map(
              (catalog) => catalogHeaders.map((key) => catalog[key]).toList()));
      final catalogProducts = <List<dynamic>>[];
      for (final catalog in catalogs) {
        final catalogId =
            catalog['slug']?.toString() ?? catalog['name']?.toString() ?? '';
        _exportImageValue(imageRows, 'logo catálogo', catalogId,
            catalog['name']?.toString() ?? '', catalog['logo']);
        for (final product in _exportMaps(catalog['products'])) {
          final image = _exportImageValue(
              imageRows,
              'producto catálogo',
              '$catalogId:${product['id'] ?? product['name']}',
              product['name']?.toString() ?? '',
              product['image']);
          catalogProducts.add([
            catalogId,
            product['id'],
            product['name'],
            product['description'],
            product['price'],
            product['stock'],
            image,
          ]);
        }
      }
      _writeExcelSheet(
          workbook,
          'Productos catálogo',
          [
            'Catálogo',
            'Producto ID',
            'Producto',
            'Descripción',
            'Precio',
            'Existencias',
            'Imagen'
          ],
          catalogProducts);
      _writeExcelSheet(
          workbook,
          'Cobros WA',
          _exportColumns(_exportMaps(remote['paymentReminders'])),
          _exportMaps(remote['paymentReminders']).map((row) =>
              _exportColumns(_exportMaps(remote['paymentReminders']))
                  .map((key) => row[key])
                  .toList()));
      _writeExcelSheet(
          workbook,
          'Imágenes',
          [
            'Tipo',
            'Registro ID',
            'Nombre',
            'Formato',
            'Parte',
            'Partes',
            'Base64 (fragmento)'
          ],
          imageRows);

      final bytes = workbook.save();
      if (bytes == null || bytes.isEmpty)
        throw Exception('Excel no pudo generar el archivo.');
      final uri = await FilePicker.saveFile(
        fileName:
            'SyncConnect_exportacion_${DateFormat('yyyyMMdd_HHmm').format(now)}.xlsx',
        bytes: Uint8List.fromList(bytes),
        dialogTitle: 'Guardar exportación de Sync Connect',
      );
      if (uri != null && mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(content: Text('Excel guardado. $remoteStatus')));
      }
    } catch (error) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(content: Text('No se pudo exportar el Excel: $error')));
      }
    } finally {
      if (mounted) setState(() => _exporting = false);
    }
  }

  Future<void> _connectWhatsApp() async {
    try {
      final r = await http
          .post(Uri.parse('$_root/api/connect'),
              headers: await _authorizedApiHeaders(json: true), body: '{}')
          .timeout(const Duration(seconds: 12));
      final d = jsonDecode(r.body) as Map<String, dynamic>;
      if (r.statusCode >= 300)
        throw Exception(d['error'] ?? 'No se pudo iniciar WhatsApp.');
      widget.onNavigate(6);
    } catch (e) {
      if (mounted) setState(() => _waError = '$e');
    }
  }

  Widget _input(TextEditingController c, String label,
          {int lines = 1, String? helper, bool secret = false}) =>
      Padding(
          padding: const EdgeInsets.only(bottom: 12),
          child: TextField(
              controller: c,
              maxLines: secret ? 1 : lines,
              obscureText: secret,
              autocorrect: !secret,
              enableSuggestions: !secret,
              decoration:
                  InputDecoration(labelText: label, helperText: helper)));
  @override
  Widget build(BuildContext context) => _loading
      ? const _Loading()
      : LayoutBuilder(builder: (context, size) {
          final compact = size.maxWidth < 760;
          final width = compact ? size.maxWidth : (size.maxWidth - 20) / 2;
          return ListView(padding: const EdgeInsets.all(24), children: [
            const _PageHeading('Configuración',
                'Personaliza el negocio y el asistente que atiende y vende por WhatsApp.'),
            if (_waError.isNotEmpty)
              Padding(
                  padding: const EdgeInsets.only(top: 12),
                  child: Text(_waError,
                      style: const TextStyle(color: Colors.red))),
            const SizedBox(height: 16),
            Wrap(
                spacing: 20,
                runSpacing: 16,
                crossAxisAlignment: WrapCrossAlignment.start,
                children: [
                  SizedBox(
                      width: width,
                      child: Card(
                          elevation: 0,
                          child: Padding(
                              padding: const EdgeInsets.all(18),
                              child: Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: [
                                    const Text('Tu negocio',
                                        style: TextStyle(
                                            fontSize: 18,
                                            fontWeight: FontWeight.w800)),
                                    const SizedBox(height: 14),
                                    _input(_name, 'Nombre del negocio'),
                                    DropdownButtonFormField<String>(
                                      value: _selectedCurrency,
                                      decoration: const InputDecoration(
                                          labelText: 'Moneda del negocio'),
                                      items: _currencyNames.entries
                                          .map((e) => DropdownMenuItem(
                                              value: e.key,
                                              child: Text(e.value)))
                                          .toList(),
                                      onChanged: (code) async {
                                        if (code == null) return;
                                        try {
                                          await widget.store
                                              .updateCurrency(code);
                                          if (mounted)
                                            setState(
                                                () => _selectedCurrency = code);
                                        } catch (error) {
                                          if (mounted)
                                            ScaffoldMessenger.of(context)
                                                .showSnackBar(SnackBar(
                                                    content: Text('$error'
                                                        .replaceFirst(
                                                            'Exception: ',
                                                            ''))));
                                        }
                                      },
                                    ),
                                    _input(
                                        _description, 'Descripción del negocio',
                                        lines: 3,
                                        helper:
                                            'El bot usará esto como contexto principal.'),
                                    _input(_goal, 'Objetivo de venta',
                                        lines: 3,
                                        helper:
                                            'Ejemplo: recomienda el producto adecuado y guía al cliente para pedirlo.'),
                                    _input(
                                        _tone, 'Tono y políticas del asistente',
                                        lines: 3),
                                    _input(_services, 'Productos o servicios',
                                        lines: 3),
                                    _input(_prices, 'Precios y condiciones',
                                        lines: 3),
                                    _input(_hours, 'Horario de atención'),
                                    _input(_address, 'Dirección o zona'),
                                    _input(_faqs,
                                        'Preguntas y respuestas frecuentes',
                                        lines: 4),
                                    _input(_quote, 'Mensaje de cotización',
                                        lines: 2),
                                    _input(_phone, 'Teléfono de contacto'),
                                    const Divider(height: 28),
                                    const Text('Datos del comprobante',
                                        style: TextStyle(
                                            fontSize: 16,
                                            fontWeight: FontWeight.w800)),
                                    const SizedBox(height: 6),
                                    const Text(
                                        'El logo se toma del diseño de tu catálogo. Estos datos aparecen en los comprobantes de ventas y compras.'),
                                    const SizedBox(height: 12),
                                    _input(
                                        _taxId, 'RUC / identificación fiscal'),
                                    _input(_receiptAddress, 'Dirección fiscal'),
                                    _input(
                                        _receiptPhone, 'Teléfono del negocio'),
                                    _input(_receiptEmail, 'Correo del negocio'),
                                    FilledButton.icon(
                                      onPressed:
                                          _savingReceipt ? null : _saveReceipt,
                                      icon: Icon(_savingReceipt
                                          ? Icons.hourglass_top
                                          : Icons.save_outlined),
                                      label: Text(_savingReceipt
                                          ? 'Guardando…'
                                          : 'Guardar datos del comprobante'),
                                    ),
                                  ])))),
                  SizedBox(
                      width: width,
                      child: Column(children: [
                        Card(
                            elevation: 0,
                            child: Padding(
                                padding: const EdgeInsets.all(18),
                                child: Column(
                                    crossAxisAlignment:
                                        CrossAxisAlignment.start,
                                    children: [
                                      const Text('Inteligencia artificial',
                                          style: TextStyle(
                                              fontSize: 18,
                                              fontWeight: FontWeight.w800)),
                                      const SizedBox(height: 6),
                                      Text(_aiSummary),
                                      const SizedBox(height: 10),
                                      if (_isAdmin)
                                        FilledButton.tonalIcon(
                                          onPressed: () async {
                                            final saved =
                                                await showDialog<bool>(
                                              context: context,
                                              builder: (_) =>
                                                  _AdminAiDialog(root: _root),
                                            );
                                            if (saved == true) await _load();
                                          },
                                          icon: const Icon(Icons
                                              .admin_panel_settings_outlined),
                                          label:
                                              const Text('Panel privado de IA'),
                                        )
                                      else
                                        const Text(
                                            'Los proveedores y sus claves los administra el propietario del sistema.',
                                            style: TextStyle(
                                                color: Colors.black54)),
                                      const SizedBox(height: 8),
                                      const Text(
                                          'El bot usa el proveedor disponible y el contexto de tu negocio para atender mensajes de WhatsApp.',
                                          style:
                                              TextStyle(color: Colors.black54))
                                    ]))),
                        Card(
                            elevation: 0,
                            child: Padding(
                                padding: const EdgeInsets.all(18),
                                child: Column(
                                    crossAxisAlignment:
                                        CrossAxisAlignment.start,
                                    children: [
                                      const Text('WhatsApp',
                                          style: TextStyle(
                                              fontSize: 18,
                                              fontWeight: FontWeight.w800)),
                                      const SizedBox(height: 8),
                                      Text(_waStatus),
                                      if (_waError.isNotEmpty)
                                        SelectableText(_waError,
                                            style: const TextStyle(
                                                color: Colors.red)),
                                      const SizedBox(height: 12),
                                      FilledButton.icon(
                                          onPressed:
                                              _saving ? null : _connectWhatsApp,
                                          icon: const Icon(Icons.qr_code_2),
                                          label: const Text('Conectar con QR')),
                                      const SizedBox(height: 6),
                                      const Text(
                                          'El QR aparece en la sección WhatsApp. El bot responderá con la configuración guardada.',
                                          style:
                                              TextStyle(color: Colors.black54))
                                    ])))
                      ]))
                ]),
            const SizedBox(height: 16),
            Card(
                elevation: 0,
                child: Padding(
                    padding: const EdgeInsets.all(18),
                    child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          const Row(children: [
                            Icon(Icons.table_view_outlined, color: _brand),
                            SizedBox(width: 12),
                            Expanded(
                                child: Text('Exportar información',
                                    style: TextStyle(
                                        fontSize: 16,
                                        fontWeight: FontWeight.w800)))
                          ]),
                          const SizedBox(height: 8),
                          const Text(
                              'Descarga inventario, clientes, ventas, cuotas, compras, gastos y datos de WhatsApp en un libro de Excel.'),
                          const SizedBox(height: 12),
                          FilledButton.icon(
                              onPressed: _exporting ? null : _exportWorkbook,
                              icon: Icon(_exporting
                                  ? Icons.hourglass_top
                                  : Icons.download_outlined),
                              label: Text(_exporting
                                  ? 'Preparando…'
                                  : 'Exportar Excel'))
                        ]))),
            const SizedBox(height: 16),
            Align(
                alignment: Alignment.centerRight,
                child: FilledButton.icon(
                    onPressed: _saving ? null : _save,
                    icon: Icon(
                        _saving ? Icons.hourglass_top : Icons.save_outlined),
                    label:
                        Text(_saving ? 'Guardando…' : 'Guardar configuración')))
          ]);
        });
}

class _AdminAiDialog extends StatefulWidget {
  const _AdminAiDialog({required this.root});
  final String root;
  @override
  State<_AdminAiDialog> createState() => _AdminAiDialogState();
}

class _AdminAiDialogState extends State<_AdminAiDialog> {
  bool _loading = true;
  bool _saving = false;
  String? _error;
  List<Map<String, dynamic>> _providers = [];
  final Map<String, TextEditingController> _models = {};
  final Map<String, TextEditingController> _keys = {};
  final Set<String> _enabled = {};
  final Set<String> _clear = {};

  @override
  void initState() {
    super.initState();
    _load();
  }

  @override
  void dispose() {
    for (final controller in [..._models.values, ..._keys.values]) {
      controller.dispose();
    }
    super.dispose();
  }

  Future<void> _load() async {
    try {
      final response = await http
          .get(
            Uri.parse('${widget.root}/api/admin/ai'),
            headers: await _authorizedApiHeaders(),
          )
          .timeout(const Duration(seconds: 20));
      final body = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300) {
        throw Exception(
            body['error'] ?? 'No se pudo abrir la administración de IA.');
      }
      final list = (body['providers'] as List?) ?? const [];
      for (final entry in list.whereType<Map>()) {
        final provider = Map<String, dynamic>.from(entry);
        final id = provider['id'].toString();
        _models[id] =
            TextEditingController(text: provider['model']?.toString() ?? '');
        _keys[id] = TextEditingController();
        if (provider['enabled'] == true) _enabled.add(id);
      }
      if (mounted)
        setState(() {
          _providers = list
              .whereType<Map>()
              .map((e) => Map<String, dynamic>.from(e))
              .toList();
          _loading = false;
        });
    } catch (error) {
      if (mounted)
        setState(() {
          _error = '$error'.replaceFirst('Exception: ', '');
          _loading = false;
        });
    }
  }

  Future<void> _save() async {
    setState(() {
      _saving = true;
      _error = null;
    });
    try {
      final response = await http
          .put(
            Uri.parse('${widget.root}/api/admin/ai'),
            headers: await _authorizedApiHeaders(json: true),
            body: jsonEncode({
              'providers': [
                for (final provider in _providers)
                  {
                    'id': provider['id'],
                    'model': _models[provider['id']]?.text.trim(),
                    'keys': _keys[provider['id']]?.text.trim() ?? '',
                    'enabled': _enabled.contains(provider['id']),
                    'clear': _clear.contains(provider['id']),
                  }
              ]
            }),
          )
          .timeout(const Duration(seconds: 25));
      final body = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300) {
        throw Exception(body['error'] ?? 'No se pudieron guardar las claves.');
      }
      if (mounted) Navigator.pop(context, true);
    } catch (error) {
      if (mounted)
        setState(() => _error = '$error'.replaceFirst('Exception: ', ''));
    } finally {
      if (mounted) setState(() => _saving = false);
    }
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Administración privada de IA'),
        content: SizedBox(
          width: 620,
          height: MediaQuery.sizeOf(context).height * .7,
          child: _loading
              ? const Center(child: CircularProgressIndicator())
              : _error != null && _providers.isEmpty
                  ? Center(child: Text(_error!, textAlign: TextAlign.center))
                  : ListView(children: [
                      const Text(
                          'Las claves se cifran en el servidor y solo este panel muestra sus estados. Deja el campo vacío para conservar las claves guardadas.'),
                      const SizedBox(height: 12),
                      if (_error != null)
                        Padding(
                            padding: const EdgeInsets.only(bottom: 8),
                            child: Text(_error!,
                                style: const TextStyle(color: Colors.red))),
                      for (final provider in _providers)
                        Card(
                          elevation: 0,
                          child: ExpansionTile(
                            leading: const Icon(Icons.auto_awesome_outlined,
                                color: _brand),
                            title: Text(provider['name']?.toString() ?? ''),
                            subtitle: Text(provider['keyCount'] == 0
                                ? 'Sin clave guardada'
                                : '${provider['keyCount']} clave(s) guardada(s)'),
                            childrenPadding:
                                const EdgeInsets.fromLTRB(16, 0, 16, 16),
                            children: [
                              SwitchListTile(
                                contentPadding: EdgeInsets.zero,
                                title: const Text('Usar este proveedor'),
                                value: _enabled.contains(provider['id']),
                                onChanged: (value) => setState(() {
                                  if (value)
                                    _enabled.add(provider['id']);
                                  else
                                    _enabled.remove(provider['id']);
                                }),
                              ),
                              TextField(
                                controller: _models[provider['id']],
                                decoration:
                                    const InputDecoration(labelText: 'Modelo'),
                              ),
                              const SizedBox(height: 10),
                              TextField(
                                controller: _keys[provider['id']],
                                obscureText: true,
                                decoration: InputDecoration(
                                  labelText: 'Clave(s) API',
                                  helperText:
                                      'Puedes pegar varias separadas por coma. No se vuelven a mostrar.',
                                ),
                              ),
                              if (_number(provider['keyCount']) > 0)
                                Align(
                                  alignment: Alignment.centerLeft,
                                  child: TextButton.icon(
                                    onPressed: () => setState(() {
                                      if (_clear.contains(provider['id']))
                                        _clear.remove(provider['id']);
                                      else
                                        _clear.add(provider['id']);
                                    }),
                                    icon: Icon(_clear.contains(provider['id'])
                                        ? Icons.undo
                                        : Icons.delete_outline),
                                    label: Text(_clear.contains(provider['id'])
                                        ? 'Conservar claves'
                                        : 'Eliminar claves guardadas'),
                                  ),
                                ),
                            ],
                          ),
                        ),
                    ]),
        ),
        actions: [
          TextButton(
              onPressed: _saving ? null : () => Navigator.pop(context),
              child: const Text('Cerrar')),
          FilledButton.icon(
            onPressed: _saving || _loading ? null : _save,
            icon: _saving
                ? const SizedBox(
                    width: 18,
                    height: 18,
                    child: CircularProgressIndicator(strokeWidth: 2))
                : const Icon(Icons.lock_outline),
            label: Text(_saving ? 'Guardando…' : 'Guardar proveedores'),
          ),
        ],
      );
}

class _CatalogPreview extends StatefulWidget {
  const _CatalogPreview({
    required this.name,
    required this.description,
    required this.logo,
    required this.primary,
    required this.accent,
    required this.theme,
    required this.currency,
    required this.products,
  });
  final String name, description, primary, accent, theme, currency;
  final String? logo;
  final List<Map<String, dynamic>> products;

  @override
  State<_CatalogPreview> createState() => _CatalogPreviewState();
}

class _CatalogPreviewState extends State<_CatalogPreview> {
  String _query = '';

  @override
  Widget build(BuildContext context) {
    final brand = _hexColor(widget.primary);
    final accent = _hexColor(widget.accent);
    final radius = switch (widget.theme) {
      'minimal' => 8.0,
      'editorial' => 4.0,
      'playful' => 24.0,
      'elegant' => 14.0,
      _ => 18.0,
    };
    final canvas = switch (widget.theme) {
      'editorial' => const Color(0xfffffbf4),
      'elegant' => const Color(0xfff3f0ea),
      'playful' => const Color(0xfff4f7ff),
      _ => const Color(0xfff6f8fb),
    };
    final shown = widget.products.where((product) {
      final query = _query.trim().toLowerCase();
      return query.isEmpty ||
          '${product['name'] ?? ''} ${product['description'] ?? ''}'
              .toLowerCase()
              .contains(query);
    }).toList();
    return Container(
      color: canvas,
      child: Column(children: [
        Container(
          color: brand,
          padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 16),
          child: Row(children: [
            if (widget.logo != null && widget.logo!.contains(',')) ...[
              ClipRRect(
                  borderRadius: BorderRadius.circular(12),
                  child: Image.memory(
                      base64Decode(widget.logo!.split(',').last),
                      width: 48,
                      height: 48,
                      fit: BoxFit.cover)),
              const SizedBox(width: 12),
            ] else ...[
              const Icon(Icons.storefront_outlined,
                  color: Colors.white, size: 32),
              const SizedBox(width: 12),
            ],
            Expanded(
                child: Text(widget.name,
                    style: const TextStyle(
                        color: Colors.white,
                        fontWeight: FontWeight.w800,
                        fontSize: 19))),
            const Icon(Icons.shopping_bag_outlined, color: Colors.white),
          ]),
        ),
        Expanded(
            child: ListView(padding: const EdgeInsets.all(22), children: [
          Center(
              child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 1080),
            child:
                Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              const Text('TIENDA EN LÍNEA',
                  style: TextStyle(
                      fontSize: 11,
                      fontWeight: FontWeight.w800,
                      letterSpacing: 2,
                      color: Colors.black54)),
              const SizedBox(height: 6),
              Text(widget.name,
                  style: TextStyle(
                      fontSize: 34, fontWeight: FontWeight.w900, color: brand)),
              if (widget.description.isNotEmpty)
                Padding(
                    padding: const EdgeInsets.only(top: 6, bottom: 18),
                    child: Text(widget.description,
                        style: const TextStyle(
                            color: Colors.black54, fontSize: 16))),
              TextField(
                decoration: InputDecoration(
                  hintText: 'Buscar productos',
                  prefixIcon: const Icon(Icons.search),
                  filled: true,
                  fillColor: Colors.white,
                  border: OutlineInputBorder(
                      borderRadius: BorderRadius.circular(radius)),
                ),
                onChanged: (value) => setState(() => _query = value),
              ),
              const SizedBox(height: 18),
              Text('${shown.length} productos',
                  style: const TextStyle(
                      fontWeight: FontWeight.w700, color: Colors.black54)),
              const SizedBox(height: 10),
              if (shown.isEmpty)
                const Padding(
                    padding: EdgeInsets.all(24),
                    child: Center(
                        child:
                            Text('No encontramos productos con esa búsqueda.')))
              else
                LayoutBuilder(builder: (context, constraints) {
                  final columns = constraints.maxWidth > 900
                      ? 4
                      : constraints.maxWidth > 600
                          ? 3
                          : constraints.maxWidth > 380
                              ? 2
                              : 1;
                  return GridView.builder(
                    itemCount: shown.length,
                    shrinkWrap: true,
                    physics: const NeverScrollableScrollPhysics(),
                    gridDelegate: SliverGridDelegateWithFixedCrossAxisCount(
                        crossAxisCount: columns,
                        crossAxisSpacing: 14,
                        mainAxisSpacing: 14,
                        childAspectRatio: .72),
                    itemBuilder: (context, index) {
                      final product = shown[index];
                      final image = product['image']?.toString() ?? '';
                      final stock = _number(product['stock']);
                      return Card(
                        elevation: widget.theme == 'minimal' ? 0 : 2,
                        color: widget.theme == 'elegant'
                            ? const Color(0xfffbfaf7)
                            : Colors.white,
                        shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(radius),
                            side: const BorderSide(color: Color(0xffe6e9ed))),
                        clipBehavior: Clip.antiAlias,
                        child: Column(
                            crossAxisAlignment: CrossAxisAlignment.stretch,
                            children: [
                              Expanded(
                                  child: image.startsWith('data:image/') &&
                                          image.contains(',')
                                      ? Image.memory(
                                          base64Decode(image.split(',').last),
                                          fit: BoxFit.cover)
                                      : Container(
                                          color: brand.withValues(alpha: .08),
                                          child: Icon(
                                              Icons.inventory_2_outlined,
                                              size: 44,
                                              color: brand))),
                              Padding(
                                  padding: const EdgeInsets.all(13),
                                  child: Column(
                                      crossAxisAlignment:
                                          CrossAxisAlignment.start,
                                      children: [
                                        Text(
                                            product['name']?.toString() ??
                                                'Producto',
                                            maxLines: 1,
                                            overflow: TextOverflow.ellipsis,
                                            style: const TextStyle(
                                                fontWeight: FontWeight.w800)),
                                        const SizedBox(height: 4),
                                        Text(
                                            product['description']
                                                    ?.toString() ??
                                                '',
                                            maxLines: 2,
                                            overflow: TextOverflow.ellipsis,
                                            style: const TextStyle(
                                                fontSize: 12,
                                                color: Colors.black54)),
                                        const SizedBox(height: 8),
                                        Row(children: [
                                          Expanded(
                                              child: Text(
                                                  NumberFormat.currency(
                                                          locale:
                                                              _currencyLocales[
                                                                  widget
                                                                      .currency],
                                                          symbol:
                                                              _currencySymbols[
                                                                  widget
                                                                      .currency],
                                                          decimalDigits: 2)
                                                      .format(_number(
                                                          product['price'])),
                                                  style: TextStyle(
                                                      fontWeight:
                                                          FontWeight.w900,
                                                      color: accent))),
                                          Text(
                                              stock > 0
                                                  ? 'Disponible'
                                                  : 'Agotado',
                                              style: TextStyle(
                                                  fontSize: 10,
                                                  color: stock > 0
                                                      ? const Color(0xff28734f)
                                                      : Colors.red)),
                                        ]),
                                      ])),
                            ]),
                      );
                    },
                  );
                }),
            ]),
          )),
        ])),
      ]),
    );
  }
}

class _CatalogPage extends StatefulWidget {
  const _CatalogPage({required this.store});
  final _Store store;
  @override
  State<_CatalogPage> createState() => _CatalogPageState();
}

class _CatalogPageState extends State<_CatalogPage> {
  static const _server = String.fromEnvironment('BOT_API_URL');
  String get _root => _resolveApiRoot(_server);
  late final _name = TextEditingController(
      text: widget.store.catalogConfig['name']?.toString() ?? '');
  late final _description = TextEditingController(
      text: widget.store.catalogConfig['description']?.toString() ?? '');
  late final _phone = TextEditingController(
      text: widget.store.catalogConfig['whatsapp']?.toString() ?? '');
  late final TextEditingController _slug;
  final Map<String, String> _productCopy = {};
  String? _logoData;
  String _primary = '#2457A7';
  String _accent = '#A16207';
  String _theme = 'modern';
  bool _saving = false;
  bool _assisting = false;
  String? _url, _error;
  @override
  void initState() {
    super.initState();
    final saved = widget.store.catalogConfig;
    _url = saved['url']?.toString();
    _slug = TextEditingController(
        text: saved['slug']?.toString() ??
            _catalogSlugFromUrl(_url) ??
            _catalogSlugFromName(_name.text));
    _logoData = saved['logo']?.toString();
    final savedCopy = saved['productDescriptions'];
    if (savedCopy is Map) {
      _productCopy.addAll(savedCopy
          .map((key, value) => MapEntry(key.toString(), value.toString())));
    }
    final color = saved['primary']?.toString() ?? '';
    if (_catalogColors.contains(color)) _primary = color;
    final accent = saved['accent']?.toString() ?? '';
    if (_catalogColors.contains(accent)) _accent = accent;
    final theme = saved['theme']?.toString() ?? '';
    if (_catalogThemes.containsValue(theme)) _theme = theme;
  }

  @override
  void dispose() {
    _name.dispose();
    _description.dispose();
    _phone.dispose();
    _slug.dispose();
    super.dispose();
  }

  Future<void> _publish() async {
    if (_root.isEmpty ||
        widget.store.products.isEmpty ||
        _name.text.trim().isEmpty ||
        _saving) return;
    setState(() {
      _saving = true;
      _error = null;
    });
    try {
      await _saveDesign();
      final response = await http
          .post(Uri.parse('$_root/api/catalogs'),
              headers: await _authorizedApiHeaders(json: true),
              body: jsonEncode({
                'name': _name.text.trim(),
                'slug': _slug.text.trim().isEmpty
                    ? _name.text.trim()
                    : _slug.text.trim(),
                'description': _description.text.trim(),
                'whatsapp': _phone.text.trim(),
                'currency': widget.store.currencyCode,
                'logo': _logoData ?? '',
                'primary': _primary,
                'accent': _accent,
                'theme': _theme,
                'productDescriptions': _productCopy,
                'products': widget.store.products
                    .map((p) => {
                          'id': p['id'],
                          'name': p['name'],
                          'description': _productCopy[p['id']?.toString()] ??
                              p['description'] ??
                              '',
                          'price': p['price'],
                          'stock': p['stock'],
                          'image': p['image'] ?? ''
                        })
                    .toList()
              }))
          .timeout(const Duration(seconds: 15));
      final data = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300)
        throw Exception(data['error'] ?? 'No se pudo publicar.');
      final publishedUrl = data['url']?.toString();
      final publishedSlug = data['slug']?.toString();
      if (publishedUrl != null) {
        await widget.store.updateCatalogConfig({
          ...widget.store.catalogConfig,
          'url': publishedUrl,
          if (publishedSlug != null) 'slug': publishedSlug,
        });
        if (mounted) {
          setState(() => _url = publishedUrl);
          if (publishedSlug != null) _slug.text = publishedSlug;
        }
      }
    } catch (e) {
      if (mounted) setState(() => _error = '$e');
    } finally {
      if (mounted) setState(() => _saving = false);
    }
  }

  Future<void> _pickLogo() async {
    try {
      final result = await FilePicker.pickFiles(
          type: FileType.image, allowMultiple: false, withData: true);
      if (result == null || result.files.isEmpty) return;
      final bytes = result.files.first.bytes;
      if (bytes == null || bytes.isEmpty)
        throw Exception('No se pudo leer el archivo.');
      if (bytes.length > 20 * 1024 * 1024)
        throw Exception('La imagen supera los 20 MB.');
      final decoded = img.decodeImage(bytes);
      if (decoded == null) throw Exception('Usa una imagen JPG, PNG o WebP.');
      final resized = img.copyResize(decoded,
          width: decoded.width >= decoded.height ? 600 : null,
          height: decoded.height > decoded.width ? 600 : null);
      var compressed = img.encodeJpg(resized, quality: 78);
      if (compressed.length > 450000) {
        compressed = img.encodeJpg(
            img.copyResize(resized,
                width: resized.width >= resized.height ? 420 : null,
                height: resized.height > resized.width ? 420 : null),
            quality: 65);
      }
      if (compressed.length > 450000)
        throw Exception('No se pudo reducir el logo. Prueba otra imagen.');
      if (mounted)
        setState(() =>
            _logoData = 'data:image/jpeg;base64,${base64Encode(compressed)}');
    } catch (e) {
      if (mounted)
        setState(() => _error = '$e'.replaceFirst('Exception: ', ''));
    }
  }

  Future<void> _saveDesign() async {
    await widget.store.updateCatalogConfig({
      'name': _name.text.trim(),
      'slug': _slug.text.trim(),
      'description': _description.text.trim(),
      'whatsapp': _phone.text.trim(),
      'logo': _logoData ?? '',
      'primary': _primary,
      'accent': _accent,
      'theme': _theme,
      'productDescriptions': _productCopy,
      if (_url != null) 'url': _url,
    });
  }

  Future<void> _assistWithAi() async {
    if (_root.isEmpty || _assisting || widget.store.products.isEmpty) return;
    setState(() {
      _assisting = true;
      _error = null;
    });
    try {
      final response = await http
          .post(
            Uri.parse('$_root/api/catalogs/assist'),
            headers: await _authorizedApiHeaders(json: true),
            body: jsonEncode({
              'name': _name.text.trim(),
              'description': _description.text.trim(),
              'products': widget.store.products
                  .map((product) => {
                        'id': product['id'],
                        'name': product['name'],
                        'description':
                            _productCopy[product['id']?.toString()] ??
                                product['description'] ??
                                '',
                      })
                  .toList(),
            }),
          )
          .timeout(const Duration(seconds: 60));
      final data = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode >= 300) {
        throw Exception(data['error'] ?? 'No se pudo pedir ayuda a la IA.');
      }
      final products = (data['products'] as List? ?? const []).whereType<Map>();
      for (final product in products) {
        final id = product['id']?.toString() ?? '';
        final description = product['description']?.toString().trim() ?? '';
        if (id.isNotEmpty && description.isNotEmpty)
          _productCopy[id] = description;
      }
      final suggestion = data['description']?.toString().trim() ?? '';
      if (suggestion.isNotEmpty) _description.text = suggestion;
      await _saveDesign();
      if (mounted)
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
              content: Text(
                  'Textos sugeridos con ${data['provider'] ?? 'IA'}. Revísalos y pulsa Publicar catálogo.')),
        );
    } catch (error) {
      if (mounted)
        setState(() => _error = '$error'.replaceFirst('Exception: ', ''));
    } finally {
      if (mounted) setState(() => _assisting = false);
    }
  }

  Future<void> _showPreview() async {
    await showDialog<void>(
      context: context,
      builder: (context) => Dialog(
        insetPadding: const EdgeInsets.all(18),
        child: SizedBox(
          width: math.min(MediaQuery.sizeOf(context).width - 24, 1050),
          height: MediaQuery.sizeOf(context).height * .9,
          child: Column(children: [
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 8),
              child: Row(children: [
                const Expanded(
                    child: Text('Vista previa de tu tienda',
                        style: TextStyle(
                            fontSize: 18, fontWeight: FontWeight.w800))),
                IconButton(
                    tooltip: 'Cerrar',
                    onPressed: () => Navigator.pop(context),
                    icon: const Icon(Icons.close)),
              ]),
            ),
            const Divider(height: 1),
            Expanded(
                child: _CatalogPreview(
              name: _name.text.trim().isEmpty ? 'Mi tienda' : _name.text.trim(),
              description: _description.text.trim(),
              logo: _logoData,
              primary: _primary,
              accent: _accent,
              theme: _theme,
              currency: widget.store.currencyCode,
              products: widget.store.products
                  .map((product) => {
                        ...product,
                        'description':
                            _productCopy[product['id']?.toString()] ??
                                product['description'] ??
                                '',
                      })
                  .toList(),
            )),
          ]),
        ),
      ),
    );
  }

  Future<void> _saveDesignInteractively() async {
    try {
      await _saveDesign();
      var published = false;
      if (_url != null &&
          _root.isNotEmpty &&
          widget.store.products.isNotEmpty) {
        await _publish();
        published = _error == null;
        if (!published) {
          if (mounted) {
            ScaffoldMessenger.of(context).showSnackBar(SnackBar(
                content: Text('No se pudo actualizar el catálogo: $_error')));
          }
          return;
        }
      }
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(SnackBar(
            content: Text(published
                ? 'Diseño guardado y catálogo actualizado.'
                : 'Diseño guardado. Pulsa “Publicar catálogo” para compartirlo.')));
      }
    } catch (e) {
      if (mounted) setState(() => _error = 'No se pudo guardar el diseño: $e');
    }
  }

  Widget _palettePicker(
          String label, String selected, ValueChanged<String> onSelect) =>
      Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
        Text(label, style: const TextStyle(color: Colors.black54)),
        const SizedBox(height: 8),
        Wrap(spacing: 10, children: [
          for (final hex in _catalogColors)
            InkWell(
              onTap: () => onSelect(hex),
              borderRadius: BorderRadius.circular(24),
              child: Container(
                width: 34,
                height: 34,
                decoration: BoxDecoration(
                  color: _hexColor(hex),
                  shape: BoxShape.circle,
                  border: Border.all(
                      color: selected == hex ? _ink : Colors.white,
                      width: selected == hex ? 3 : 1),
                ),
              ),
            )
        ]),
      ]);

  @override
  Widget build(BuildContext context) =>
      LayoutBuilder(builder: (context, constraints) {
        final compact = constraints.maxWidth < 680;
        return ListView(padding: const EdgeInsets.all(24), children: [
          const _PageHeading('Catálogo web',
              'Personaliza tu marca, edita la presentación y publica tus productos.'),
          const SizedBox(height: 18),
          Wrap(
              spacing: 18,
              runSpacing: 18,
              crossAxisAlignment: WrapCrossAlignment.start,
              children: [
                SizedBox(
                    width: compact
                        ? constraints.maxWidth
                        : (constraints.maxWidth - 18) / 2,
                    child: Card(
                        elevation: 0,
                        child: Padding(
                            padding: const EdgeInsets.all(20),
                            child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  const Text('Información del negocio',
                                      style: TextStyle(
                                          fontSize: 18,
                                          fontWeight: FontWeight.w800)),
                                  const SizedBox(height: 14),
                                  const Text('Diseño de marca',
                                      style: TextStyle(
                                          fontWeight: FontWeight.w700)),
                                  const SizedBox(height: 10),
                                  DropdownButtonFormField<String>(
                                    initialValue: _theme,
                                    decoration: const InputDecoration(
                                        labelText: 'Estilo de página'),
                                    items: _catalogThemes.entries
                                        .map((entry) => DropdownMenuItem(
                                            value: entry.value,
                                            child: Text(entry.key)))
                                        .toList(),
                                    onChanged: (value) {
                                      if (value != null)
                                        setState(() => _theme = value);
                                    },
                                  ),
                                  const SizedBox(height: 12),
                                  Row(children: [
                                    _logoData != null &&
                                            _logoData!.contains(',')
                                        ? ClipRRect(
                                            borderRadius:
                                                BorderRadius.circular(14),
                                            child: Image.memory(
                                                base64Decode(
                                                    _logoData!.split(',').last),
                                                width: 64,
                                                height: 64,
                                                fit: BoxFit.cover))
                                        : CircleAvatar(
                                            radius: 32,
                                            backgroundColor: Color(0xffedf3ff),
                                            child: Icon(
                                                Icons.storefront_outlined,
                                                color: _hexColor(_primary))),
                                    const SizedBox(width: 12),
                                    Expanded(
                                        child: Wrap(
                                            spacing: 8,
                                            runSpacing: 8,
                                            children: [
                                          OutlinedButton.icon(
                                              onPressed: _pickLogo,
                                              icon: const Icon(
                                                  Icons.upload_outlined),
                                              label: Text(_logoData == null
                                                  ? 'Subir logo'
                                                  : 'Cambiar logo')),
                                          if (_logoData != null)
                                            TextButton(
                                                onPressed: () => setState(
                                                    () => _logoData = null),
                                                child: const Text('Quitar')),
                                        ])),
                                  ]),
                                  const SizedBox(height: 12),
                                  _palettePicker(
                                      'Color principal',
                                      _primary,
                                      (value) =>
                                          setState(() => _primary = value)),
                                  const SizedBox(height: 12),
                                  _palettePicker(
                                      'Color de acento',
                                      _accent,
                                      (value) =>
                                          setState(() => _accent = value)),
                                  const SizedBox(height: 8),
                                  Align(
                                      alignment: Alignment.centerRight,
                                      child: TextButton.icon(
                                          onPressed: _saveDesignInteractively,
                                          icon: const Icon(Icons.save_outlined),
                                          label: const Text('Guardar diseño'))),
                                  const Divider(height: 24),
                                  TextField(
                                      controller: _name,
                                      decoration: const InputDecoration(
                                          labelText: 'Nombre del negocio')),
                                  const SizedBox(height: 12),
                                  TextField(
                                      controller: _slug,
                                      keyboardType: TextInputType.url,
                                      autocorrect: false,
                                      decoration: const InputDecoration(
                                          labelText:
                                              'Nombre para el enlace público',
                                          hintText: 'mi-tienda',
                                          helperText:
                                              'Usa letras, números y guiones; por ejemplo, mi-tienda.')),
                                  const SizedBox(height: 12),
                                  TextField(
                                      controller: _description,
                                      maxLines: 3,
                                      decoration: const InputDecoration(
                                          labelText: 'Descripción',
                                          hintText: 'Cuéntales qué ofreces')),
                                  const SizedBox(height: 10),
                                  Wrap(spacing: 8, runSpacing: 8, children: [
                                    OutlinedButton.icon(
                                      onPressed: _showPreview,
                                      icon:
                                          const Icon(Icons.visibility_outlined),
                                      label: const Text('Vista previa'),
                                    ),
                                    FilledButton.tonalIcon(
                                      onPressed: _assisting ||
                                              widget.store.products.isEmpty
                                          ? null
                                          : _assistWithAi,
                                      icon: _assisting
                                          ? const SizedBox(
                                              width: 18,
                                              height: 18,
                                              child: CircularProgressIndicator(
                                                  strokeWidth: 2))
                                          : const Icon(
                                              Icons.auto_awesome_outlined),
                                      label: Text(_assisting
                                          ? 'Generando…'
                                          : 'Ayuda con IA'),
                                    ),
                                  ]),
                                  const Padding(
                                    padding: EdgeInsets.only(top: 6),
                                    child: Text(
                                        'La IA sugiere descripciones con los datos de tus productos; revísalas antes de publicar.',
                                        style: TextStyle(
                                            color: Colors.black54,
                                            fontSize: 12)),
                                  ),
                                  const SizedBox(height: 12),
                                  TextField(
                                      controller: _phone,
                                      keyboardType: TextInputType.phone,
                                      decoration: const InputDecoration(
                                          labelText: 'WhatsApp de contacto',
                                          hintText: 'Código de país y número')),
                                  const SizedBox(height: 16),
                                  FilledButton.icon(
                                      onPressed: _saving ||
                                              widget.store.products.isEmpty
                                          ? null
                                          : _publish,
                                      icon: Icon(_saving
                                          ? Icons.hourglass_top
                                          : Icons.language),
                                      label: Text(_saving
                                          ? 'Publicando…'
                                          : 'Publicar catálogo')),
                                  if (_error != null)
                                    Padding(
                                        padding: const EdgeInsets.only(top: 12),
                                        child: Text(_error!,
                                            style: const TextStyle(
                                                color: Colors.red))),
                                  if (_url != null) ...[
                                    const SizedBox(height: 16),
                                    const Text('Página publicada',
                                        style: TextStyle(
                                            fontWeight: FontWeight.w700)),
                                    SelectableText(_url!,
                                        style: const TextStyle(color: _brand)),
                                    Align(
                                        alignment: Alignment.centerRight,
                                        child: TextButton.icon(
                                            onPressed: () async {
                                              await Clipboard.setData(
                                                  ClipboardData(text: _url!));
                                              if (context.mounted)
                                                ScaffoldMessenger.of(context)
                                                    .showSnackBar(const SnackBar(
                                                        content: Text(
                                                            'Enlace copiado')));
                                            },
                                            icon: const Icon(Icons.copy),
                                            label: const Text('Copiar enlace')))
                                  ]
                                ])))),
                SizedBox(
                    width: compact
                        ? constraints.maxWidth
                        : (constraints.maxWidth - 18) / 2,
                    child: Card(
                        elevation: 0,
                        child: Padding(
                            padding: const EdgeInsets.all(18),
                            child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  Text(
                                      'Productos del inventario · ${widget.store.products.length}',
                                      style: const TextStyle(
                                          fontSize: 18,
                                          fontWeight: FontWeight.w800)),
                                  const SizedBox(height: 10),
                                  if (widget.store.products.isEmpty)
                                    const Padding(
                                        padding:
                                            EdgeInsets.symmetric(vertical: 20),
                                        child: Text(
                                            'Agrega productos en Inventario para incluirlos en tu catálogo.'))
                                  else
                                    for (final p in widget.store.products)
                                      ListTile(
                                          contentPadding: EdgeInsets.zero,
                                          leading:
                                              _productThumbnail(p, size: 44),
                                          title:
                                              Text(p['name']?.toString() ?? ''),
                                          subtitle: Text(
                                              'Disponible: ${_number(p['stock'])}'),
                                          trailing: Text(
                                              _money(_number(p['price'])),
                                              style: const TextStyle(
                                                  fontWeight: FontWeight.w700)))
                                ]))))
              ]),
          const SizedBox(height: 14),
          const Text(
              'El enlace usa este servidor. Para compartirlo fuera de tu red, publica el servidor en un dominio HTTPS y configura PUBLIC_BASE_URL.',
              style: TextStyle(color: Colors.black54))
        ]);
      });
}

Color _hexColor(String value) {
  final normalized = value.replaceFirst('#', '');
  if (!RegExp(r'^[0-9A-Fa-f]{6}$').hasMatch(normalized)) return _brand;
  return Color(int.parse('ff$normalized', radix: 16));
}

Widget _productThumbnail(Map<String, dynamic> product, {double size = 48}) {
  final image = product['image']?.toString() ?? '';
  if (image.startsWith('data:image/') && image.contains(',')) {
    try {
      return ClipRRect(
        borderRadius: BorderRadius.circular(12),
        child: Image.memory(base64Decode(image.split(',').last),
            width: size, height: size, fit: BoxFit.cover),
      );
    } catch (_) {}
  }
  return CircleAvatar(
    radius: size / 2,
    backgroundColor: const Color(0xffedf3ff),
    child: Icon(Icons.inventory_2_outlined, color: _brand, size: size * .52),
  );
}

class _WhatsAppPage extends StatefulWidget {
  const _WhatsAppPage();

  @override
  State<_WhatsAppPage> createState() => _WhatsAppPageState();
}

class _WhatsAppPageState extends State<_WhatsAppPage> {
  static const _configuredUrl = String.fromEnvironment('BOT_API_URL');
  String get _apiRoot => _resolveApiRoot(_configuredUrl);

  Timer? _poller;
  bool _loading = true;
  bool _fetching = false;
  bool _sending = false;
  String? _error;
  String _connectionError = '';
  String _status = 'Buscando el servidor de WhatsApp…';
  String? _qr;
  List<Map<String, dynamic>> _chats = [];
  String? _selectedJid;
  final _messageController = TextEditingController();

  @override
  void initState() {
    super.initState();
    _refresh();
    _poller = Timer.periodic(const Duration(seconds: 3), (_) => _refresh());
  }

  @override
  void dispose() {
    _poller?.cancel();
    _messageController.dispose();
    super.dispose();
  }

  Future<void> _refresh() async {
    if (_fetching) return;
    if (_apiRoot.isEmpty) {
      setState(() {
        _loading = false;
        _error = 'Configura BOT_API_URL con la dirección HTTPS de tu servidor.';
        _status = 'Servidor sin configurar';
      });
      return;
    }
    _fetching = true;
    try {
      final response = await http
          .get(Uri.parse('$_apiRoot/api/state'),
              headers: await _authorizedApiHeaders())
          .timeout(const Duration(seconds: 5));
      if (response.statusCode != 200) {
        throw Exception('El servidor respondió HTTP ${response.statusCode}.');
      }
      final state = jsonDecode(response.body) as Map<String, dynamic>;
      final chats = _toMapList(state['chats']);
      if (!mounted) return;
      setState(() {
        _chats = chats;
        _status = state['waStatus']?.toString() ?? 'Estado desconocido';
        _connectionError = state['waLastError']?.toString() ?? '';
        _qr = state['currentQr']?.toString();
        _error = null;
        if (_selectedJid == null ||
            !_chats.any((chat) => chat['id'] == _selectedJid)) {
          _selectedJid = _chats.isEmpty ? null : _chats.first['id'] as String?;
        }
      });
    } catch (error) {
      if (mounted) {
        setState(() {
          _error = 'No se pudo acceder al servidor local ($error).';
          _status = 'Servidor de WhatsApp no disponible';
          _qr = null;
        });
      }
    } finally {
      _fetching = false;
      if (mounted && _loading) setState(() => _loading = false);
    }
  }

  Future<void> _control(String path) async {
    try {
      final response = await http
          .post(Uri.parse('$_apiRoot$path'),
              headers: await _authorizedApiHeaders(json: true), body: '{}')
          .timeout(const Duration(seconds: 10));
      final result = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode < 200 || response.statusCode >= 300) {
        throw Exception(result['error'] ?? 'No se pudo completar la acción.');
      }
      await _refresh();
    } catch (error) {
      if (mounted) {
        ScaffoldMessenger.of(context)
            .showSnackBar(SnackBar(content: Text('$error')));
      }
    }
  }

  Future<void> _send() async {
    final jid = _selectedJid;
    final text = _messageController.text.trim();
    if (jid == null || text.isEmpty || _sending) return;
    setState(() => _sending = true);
    try {
      final response = await http
          .post(
              Uri.parse(
                  '$_apiRoot/api/chats/${Uri.encodeComponent(jid)}/messages'),
              headers: await _authorizedApiHeaders(json: true),
              body: jsonEncode({'text': text}))
          .timeout(const Duration(seconds: 15));
      final result = jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode < 200 || response.statusCode >= 300) {
        throw Exception(result['error'] ?? 'No se pudo enviar el mensaje.');
      }
      _messageController.clear();
      await _refresh();
    } catch (error) {
      if (mounted) {
        ScaffoldMessenger.of(context)
            .showSnackBar(SnackBar(content: Text('$error')));
      }
    } finally {
      if (mounted) setState(() => _sending = false);
    }
  }

  @override
  Widget build(BuildContext context) {
    final connected = _status.toLowerCase().contains('conectado') &&
        !_status.toLowerCase().contains('desconectado');
    final selected = _chats.where((chat) => chat['id'] == _selectedJid);
    final chat = selected.isEmpty ? null : selected.first;
    final qrData = _qr;
    return Padding(
      padding: const EdgeInsets.all(24),
      child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
        const _PageHeading('WhatsApp',
            'Vincula el teléfono, revisa chats y responde desde el CRM.'),
        const SizedBox(height: 14),
        Card(
          elevation: 0,
          child: Padding(
            padding: const EdgeInsets.all(14),
            child: Row(children: [
              Icon(connected ? Icons.check_circle : Icons.link,
                  color: connected ? _brand : Colors.orange.shade700),
              const SizedBox(width: 10),
              Expanded(
                  child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                    Text(_status,
                        style: const TextStyle(fontWeight: FontWeight.w700)),
                    Text(_error ?? '${_chats.length} conversaciones',
                        maxLines: 2,
                        overflow: TextOverflow.ellipsis,
                        style: const TextStyle(
                            color: Colors.black54, fontSize: 12)),
                    if (_error == null && _connectionError.isNotEmpty)
                      Text(_connectionError,
                          maxLines: 2,
                          overflow: TextOverflow.ellipsis,
                          style:
                              const TextStyle(color: Colors.red, fontSize: 12)),
                  ])),
              if (!connected)
                FilledButton.tonalIcon(
                    onPressed: _apiRoot.isEmpty || _loading || _error != null
                        ? null
                        : () => _control('/api/connect'),
                    icon: const Icon(Icons.qr_code_2),
                    label: const Text('Conectar'))
              else
                TextButton.icon(
                    onPressed: () async {
                      final close = await showDialog<bool>(
                          context: context,
                          builder: (context) => AlertDialog(
                                title: const Text('Cerrar WhatsApp'),
                                content: const Text(
                                    'Se desvinculará esta sesión del servidor local.'),
                                actions: [
                                  TextButton(
                                      onPressed: () =>
                                          Navigator.pop(context, false),
                                      child: const Text('Cancelar')),
                                  FilledButton(
                                      onPressed: () =>
                                          Navigator.pop(context, true),
                                      child: const Text('Cerrar sesión')),
                                ],
                              ));
                      if (close == true) await _control('/api/logout');
                    },
                    icon: const Icon(Icons.logout),
                    label: const Text('Desconectar')),
            ]),
          ),
        ),
        if (qrData != null && qrData.contains(','))
          Card(
            elevation: 0,
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Row(children: [
                Image.memory(base64Decode(qrData.split(',').last),
                    width: 120, height: 120),
                const SizedBox(width: 12),
                const Expanded(
                    child: Text(
                        'En WhatsApp, abre Dispositivos vinculados → Vincular un dispositivo y escanea este código.')),
              ]),
            ),
          ),
        const SizedBox(height: 10),
        if (_loading)
          const Expanded(child: Center(child: CircularProgressIndicator()))
        else if (_error != null)
          Expanded(
              child: _Empty(
                  icon: Icons.link_off,
                  text:
                      'Inicia el servidor con Iniciar-CRM.bat en este computador. La bandeja usa el servidor local para enlazar el QR y los mensajes.'))
        else
          Expanded(child: LayoutBuilder(builder: (context, constraints) {
            if (constraints.maxWidth < 680) {
              return Column(children: [
                DropdownButtonFormField<String>(
                  initialValue: _selectedJid,
                  decoration: const InputDecoration(labelText: 'Conversación'),
                  items: _chats
                      .map((item) => DropdownMenuItem<String>(
                          value: item['id'] as String?,
                          child: Text(item['name']?.toString() ??
                              item['id']?.toString() ??
                              'Chat')))
                      .toList(),
                  onChanged: (value) => setState(() => _selectedJid = value),
                ),
                const SizedBox(height: 8),
                Expanded(child: _thread(chat, connected)),
              ]);
            }
            return Row(children: [
              SizedBox(width: 260, child: _chatList()),
              const SizedBox(width: 12),
              Expanded(child: _thread(chat, connected)),
            ]);
          })),
      ]),
    );
  }

  Widget _chatList() => Card(
        elevation: 0,
        child: Column(children: [
          const ListTile(
              leading: Icon(Icons.forum_outlined, color: _brand),
              title: Text('Conversaciones')),
          const Divider(height: 1),
          Expanded(
              child: _chats.isEmpty
                  ? const Center(
                      child: Padding(
                          padding: EdgeInsets.all(18),
                          child: Text(
                              'Los chats aparecerán al recibir mensajes.',
                              textAlign: TextAlign.center)))
                  : ListView.builder(
                      itemCount: _chats.length,
                      itemBuilder: (context, index) {
                        final item = _chats[index];
                        final id = item['id'] as String?;
                        final messages = _toMapList(item['messages']);
                        final last = messages.isEmpty
                            ? ''
                            : messages.last['text']?.toString() ?? '';
                        return ListTile(
                          selected: id == _selectedJid,
                          selectedTileColor: const Color(0xffedf3ff),
                          title: Text(item['name']?.toString() ?? id ?? 'Chat',
                              maxLines: 1, overflow: TextOverflow.ellipsis),
                          subtitle: Text(last,
                              maxLines: 1, overflow: TextOverflow.ellipsis),
                          onTap: id == null
                              ? null
                              : () => setState(() => _selectedJid = id),
                        );
                      },
                    )),
        ]),
      );

  Widget _thread(Map<String, dynamic>? chat, bool connected) => Card(
        elevation: 0,
        child: Column(children: [
          ListTile(
              leading: const CircleAvatar(
                  backgroundColor: Color(0xffedf3ff),
                  child: Icon(Icons.person_outline, color: _brand)),
              title: Text(chat?['name']?.toString() ?? 'Selecciona un chat',
                  style: const TextStyle(fontWeight: FontWeight.w700)),
              subtitle: Text(chat?['id']?.toString() ??
                  (connected
                      ? 'Esperando conversaciones'
                      : 'Conecta WhatsApp para responder'))),
          const Divider(height: 1),
          Expanded(
              child: chat == null
                  ? const _Empty(
                      icon: Icons.chat_bubble_outline,
                      text: 'Cuando un cliente escriba, podrás responder aquí.')
                  : ListView(
                      padding: const EdgeInsets.all(12),
                      children: _toMapList(chat['messages']).map((message) {
                        final mine = message['from'] != 'customer';
                        return Align(
                          alignment: mine
                              ? Alignment.centerRight
                              : Alignment.centerLeft,
                          child: Container(
                            constraints: const BoxConstraints(maxWidth: 500),
                            margin: const EdgeInsets.symmetric(vertical: 4),
                            padding: const EdgeInsets.symmetric(
                                horizontal: 12, vertical: 9),
                            decoration: BoxDecoration(
                                color: mine
                                    ? const Color(0xffe6eeff)
                                    : const Color(0xfff0f2f0),
                                borderRadius: BorderRadius.circular(14)),
                            child: Text(message['text']?.toString() ?? ''),
                          ),
                        );
                      }).toList(),
                    )),
          Padding(
              padding: const EdgeInsets.all(10),
              child: Row(children: [
                Expanded(
                    child: TextField(
                  controller: _messageController,
                  enabled: connected && chat != null && !_sending,
                  onSubmitted: (_) => _send(),
                  decoration: const InputDecoration(
                      hintText: 'Escribe una respuesta…',
                      contentPadding:
                          EdgeInsets.symmetric(horizontal: 14, vertical: 10)),
                )),
                const SizedBox(width: 8),
                IconButton.filled(
                    tooltip: 'Enviar mensaje',
                    onPressed:
                        connected && chat != null && !_sending ? _send : null,
                    icon: _sending
                        ? const SizedBox(
                            width: 18,
                            height: 18,
                            child: CircularProgressIndicator(strokeWidth: 2))
                        : const Icon(Icons.send)),
              ])),
        ]),
      );
}

List<Map<String, dynamic>> _toMapList(dynamic value) => value is List
    ? value
        .whereType<Map>()
        .map((item) => Map<String, dynamic>.from(item))
        .toList()
    : <Map<String, dynamic>>[];

class _ProductDialog extends StatefulWidget {
  const _ProductDialog({this.initial});
  final Map<String, dynamic>? initial;
  @override
  State<_ProductDialog> createState() => _ProductDialogState();
}

class _ProductDialogState extends State<_ProductDialog> {
  final _form = GlobalKey<FormState>();
  late final _name = TextEditingController(text: widget.initial?['name'] ?? '');
  late final _sku = TextEditingController(text: widget.initial?['sku'] ?? '');
  late final _price =
      TextEditingController(text: '${widget.initial?['price'] ?? ''}');
  late final _cost =
      TextEditingController(text: '${widget.initial?['cost'] ?? ''}');
  late final _stock =
      TextEditingController(text: '${widget.initial?['stock'] ?? '0'}');
  late final _min =
      TextEditingController(text: '${widget.initial?['minStock'] ?? '0'}');
  String? _imageData;
  String? _imageError;
  bool _loadingImage = false;
  @override
  void initState() {
    super.initState();
    _imageData = widget.initial?['image']?.toString();
  }

  Future<void> _chooseImage() async {
    setState(() {
      _loadingImage = true;
      _imageError = null;
    });
    try {
      final selection = await FilePicker.pickFiles(
          type: FileType.image, allowMultiple: false, withData: true);
      if (selection == null || selection.files.isEmpty) return;
      final bytes = selection.files.first.bytes;
      if (bytes == null || bytes.isEmpty)
        throw Exception('No se pudo leer la imagen.');
      if (bytes.length > 20 * 1024 * 1024)
        throw Exception(
            'La imagen supera los 20 MB. Elige un archivo más pequeño.');
      final decoded = img.decodeImage(bytes);
      if (decoded == null)
        throw Exception(
            'Ese formato de imagen no es compatible. Usa JPG, PNG o WebP.');
      final longestSide = math.max(decoded.width, decoded.height);
      final resized = longestSide > 1000
          ? img.copyResize(decoded,
              width: decoded.width >= decoded.height ? 1000 : null,
              height: decoded.height > decoded.width ? 1000 : null)
          : decoded;
      var quality = 82;
      var compressed = img.encodeJpg(resized, quality: quality);
      while (compressed.length > 350000 && quality > 50) {
        quality -= 8;
        compressed = img.encodeJpg(resized, quality: quality);
      }
      if (compressed.length > 350000) {
        final smaller = img.copyResize(resized,
            width: resized.width >= resized.height ? 700 : null,
            height: resized.height > resized.width ? 700 : null);
        compressed = img.encodeJpg(smaller, quality: 64);
      }
      if (compressed.length > 350000)
        throw Exception('No se pudo optimizar la imagen. Elige otra imagen.');
      if (mounted)
        setState(() =>
            _imageData = 'data:image/jpeg;base64,${base64Encode(compressed)}');
    } catch (e) {
      if (mounted)
        setState(() => _imageError = '$e'.replaceFirst('Exception: ', ''));
    } finally {
      if (mounted) setState(() => _loadingImage = false);
    }
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
          title: Text(
              widget.initial == null ? 'Nuevo producto' : 'Editar producto'),
          content: SizedBox(
              width: 440,
              child: Form(
                  key: _form,
                  child: SingleChildScrollView(
                      child: Column(mainAxisSize: MainAxisSize.min, children: [
                    _field(_name, 'Nombre del producto', required: true),
                    _field(_sku, 'Código / SKU'),
                    Row(children: [
                      if (_imageData != null &&
                          _imageData!.startsWith('data:image/'))
                        ClipRRect(
                            borderRadius: BorderRadius.circular(12),
                            child: Image.memory(
                                base64Decode(_imageData!.split(',').last),
                                width: 72,
                                height: 72,
                                fit: BoxFit.cover))
                      else
                        const CircleAvatar(
                            radius: 36,
                            backgroundColor: Color(0xffedf3ff),
                            child: Icon(Icons.image_outlined,
                                color: _brand, size: 28)),
                      const SizedBox(width: 12),
                      Expanded(
                          child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                            OutlinedButton.icon(
                                onPressed: _loadingImage ? null : _chooseImage,
                                icon: Icon(_loadingImage
                                    ? Icons.hourglass_top
                                    : Icons.add_photo_alternate_outlined),
                                label: Text(_loadingImage
                                    ? 'Procesando imagen…'
                                    : _imageData == null
                                        ? 'Subir imagen'
                                        : 'Cambiar imagen')),
                            if (_imageData != null)
                              TextButton.icon(
                                  onPressed: () =>
                                      setState(() => _imageData = null),
                                  icon: const Icon(Icons.delete_outline,
                                      size: 18),
                                  label: const Text('Quitar imagen'))
                          ]))
                    ]),
                    if (_imageError != null)
                      Align(
                          alignment: Alignment.centerLeft,
                          child: Text(_imageError!,
                              style: const TextStyle(color: Colors.red))),
                    const SizedBox(height: 8),
                    Row(children: [
                      Expanded(
                          child: _field(_price, 'Precio de venta',
                              number: true, required: true)),
                      const SizedBox(width: 10),
                      Expanded(child: _field(_cost, 'Costo', number: true))
                    ]),
                    Row(children: [
                      Expanded(
                          child: _field(_stock, 'Existencias', number: true)),
                      const SizedBox(width: 10),
                      Expanded(
                          child: _field(_min, 'Stock mínimo', number: true))
                    ]),
                  ])))),
          actions: [
            TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('Cancelar')),
            FilledButton(
                onPressed: () {
                  if (!_form.currentState!.validate()) return;
                  Navigator.pop(context, {
                    'name': _name.text.trim(),
                    'sku': _sku.text.trim(),
                    'price': _number(_price.text),
                    'cost': _number(_cost.text),
                    'stock': _number(_stock.text),
                    'minStock': _number(_min.text),
                    'image': _imageData ?? ''
                  });
                },
                child: const Text('Guardar'))
          ]);
}

class _ContactDialog extends StatefulWidget {
  const _ContactDialog({this.initial});
  final Map<String, dynamic>? initial;
  @override
  State<_ContactDialog> createState() => _ContactDialogState();
}

class _ContactDialogState extends State<_ContactDialog> {
  late final _name = TextEditingController(text: widget.initial?['name'] ?? '');
  late final _phone =
      TextEditingController(text: widget.initial?['phone'] ?? '');
  late final _email =
      TextEditingController(text: widget.initial?['email'] ?? '');
  late final _notes =
      TextEditingController(text: widget.initial?['notes'] ?? '');
  @override
  Widget build(BuildContext context) => AlertDialog(
          title:
              Text(widget.initial == null ? 'Nuevo cliente' : 'Editar cliente'),
          content: SizedBox(
              width: 430,
              child: Column(mainAxisSize: MainAxisSize.min, children: [
                _field(_name, 'Nombre', required: true),
                _field(_phone, 'Teléfono'),
                _field(_email, 'Correo'),
                _field(_notes, 'Notas', lines: 2)
              ])),
          actions: [
            TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('Cancelar')),
            FilledButton(
                onPressed: () {
                  if (_name.text.trim().isEmpty) return;
                  Navigator.pop(context, {
                    'name': _name.text.trim(),
                    'phone': _phone.text.trim(),
                    'email': _email.text.trim(),
                    'notes': _notes.text.trim()
                  });
                },
                child: const Text('Guardar'))
          ]);
}

class _SaleReturnDialog extends StatefulWidget {
  const _SaleReturnDialog({required this.sale});
  final Map<String, dynamic> sale;

  @override
  State<_SaleReturnDialog> createState() => _SaleReturnDialogState();
}

class _SaleReturnDialogState extends State<_SaleReturnDialog> {
  late final List<_SaleReturnLine> _lines = _makeLines();
  final _reason = TextEditingController();
  String _refundMethod = 'cash';

  List<_SaleReturnLine> _makeLines() {
    final combined = <String, Map<String, dynamic>>{};
    final raw = widget.sale['items'];
    if (raw is List) {
      for (final item in raw.whereType<Map>()) {
        final id = item['productId']?.toString() ?? '';
        final quantity = _number(item['quantity']);
        if (id.isEmpty || quantity <= 0) continue;
        final previous = combined[id];
        if (previous == null) {
          combined[id] = {
            'productId': id,
            'name': item['name']?.toString() ?? 'Producto',
            'quantity': quantity,
            'total': _number(item['total']),
          };
        } else {
          previous['quantity'] = _number(previous['quantity']) + quantity;
          previous['total'] =
              _number(previous['total']) + _number(item['total']);
        }
      }
    }
    final returnedRaw = widget.sale['returnedQuantities'];
    final returned = returnedRaw is Map
        ? returnedRaw
            .map((key, value) => MapEntry(key.toString(), _number(value)))
        : <String, double>{};
    return combined.values
        .map((item) {
          final quantity = _number(item['quantity']);
          final unitPrice =
              quantity > 0 ? _number(item['total']) / quantity : 0.0;
          final id = item['productId'].toString();
          return _SaleReturnLine(
            productId: id,
            name: item['name']?.toString() ?? 'Producto',
            unitPrice: unitPrice,
            soldQuantity: quantity,
            returnedQuantity: returned[id] ?? 0,
          );
        })
        .where((line) => line.availableQuantity > .000001)
        .toList();
  }

  @override
  void dispose() {
    for (final line in _lines) {
      line.quantity.dispose();
    }
    _reason.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => StatefulBuilder(
        builder: (context, refresh) {
          final selected =
              _lines.where((line) => _number(line.quantity.text) > 0).toList();
          final invalidQuantity = _lines.any((line) {
            final value = _number(line.quantity.text);
            return value < 0 || value > line.availableQuantity + .000001;
          });
          final subtotal = _roundMoney(selected.fold<double>(
              0,
              (sum, line) =>
                  sum + _number(line.quantity.text) * line.unitPrice));
          final originalSubtotal = _number(widget.sale['subtotal']) > 0
              ? _number(widget.sale['subtotal'])
              : _number(widget.sale['total']) -
                  _number(widget.sale['pocketSurcharge']);
          final oldReturnedSubtotal =
              _number(widget.sale['returnedItemsSubtotal']);
          final originalFee = _number(widget.sale['pocketSurcharge']);
          final newFeeTotal = originalFee > 0 && originalSubtotal > 0
              ? _roundMoney(originalFee *
                  ((oldReturnedSubtotal + subtotal) / originalSubtotal))
              : 0.0;
          final feeForThisReturn = _roundMoney(
              newFeeTotal - _number(widget.sale['returnedPocketSurcharge']));
          final returnTotal = _roundMoney(subtotal + feeForThisReturn);
          final balanceAdjusted = _roundMoney(
              math.min(_number(widget.sale['balance']), returnTotal));
          final refundDue = _roundMoney(returnTotal - balanceAdjusted);
          return AlertDialog(
            title: Text('Devolución · ${widget.sale['number'] ?? ''}'),
            content: SizedBox(
              width: 500,
              height: math.min(MediaQuery.sizeOf(context).height * .68, 560),
              child: ListView(
                children: [
                  const Text(
                      'Indica qué productos y cantidades regresó el cliente.'),
                  const SizedBox(height: 12),
                  for (final line in _lines)
                    Card(
                      elevation: 0,
                      color: const Color(0xfff4f7fc),
                      child: Padding(
                        padding: const EdgeInsets.all(12),
                        child: Row(children: [
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(line.name,
                                    style: const TextStyle(
                                        fontWeight: FontWeight.w700)),
                                const SizedBox(height: 4),
                                Text(
                                    'Devueltos antes: ${line.returnedQuantity} · Disponibles: ${line.availableQuantity} · ${_money(line.unitPrice)} c/u',
                                    style: const TextStyle(
                                        fontSize: 12, color: Colors.black54)),
                              ],
                            ),
                          ),
                          const SizedBox(width: 12),
                          SizedBox(
                            width: 100,
                            child: TextField(
                              controller: line.quantity,
                              keyboardType:
                                  const TextInputType.numberWithOptions(
                                      decimal: true),
                              decoration: const InputDecoration(
                                labelText: 'Cantidad',
                                isDense: true,
                              ),
                              onChanged: (_) => refresh(() {}),
                            ),
                          ),
                        ]),
                      ),
                    ),
                  TextField(
                    controller: _reason,
                    maxLines: 2,
                    decoration: const InputDecoration(
                      labelText: 'Motivo (opcional)',
                      hintText: 'Ej. producto defectuoso o cambio',
                    ),
                  ),
                  const SizedBox(height: 12),
                  if (invalidQuantity)
                    const Text(
                        'La cantidad debe estar entre cero y las unidades disponibles.',
                        style: TextStyle(color: Colors.red)),
                  if (selected.isNotEmpty && !invalidQuantity) ...[
                    _returnPreviewRow('Productos', _money(subtotal)),
                    if (feeForThisReturn > 0)
                      _returnPreviewRow('Parte del recargo Pocket (5%)',
                          _money(feeForThisReturn)),
                    _returnPreviewRow(
                        'Total de devolución', _money(returnTotal),
                        strong: true),
                    if (balanceAdjusted > 0)
                      _returnPreviewRow('Aplicado al saldo pendiente',
                          _money(balanceAdjusted)),
                    if (refundDue > 0) ...[
                      _returnPreviewRow(
                          'Reembolso al cliente', _money(refundDue)),
                      DropdownButtonFormField<String>(
                        value: _refundMethod,
                        decoration: const InputDecoration(
                            labelText: 'Forma de reembolso'),
                        items: const [
                          DropdownMenuItem(
                              value: 'cash', child: Text('Efectivo')),
                          DropdownMenuItem(
                              value: 'transfer', child: Text('Transferencia')),
                          DropdownMenuItem(
                              value: 'pocket_external',
                              child: Text(
                                  'Pocket · ya procesado fuera de la app')),
                        ],
                        onChanged: (value) => refresh(() {
                          if (value != null) _refundMethod = value;
                        }),
                      ),
                      if (_refundMethod == 'pocket_external')
                        const Padding(
                          padding: EdgeInsets.only(top: 6),
                          child: Text(
                              'Sync Connect no procesa reembolsos de Poket. Confirma el reembolso en Poket antes de guardar.',
                              style: TextStyle(
                                  color: Colors.black54, fontSize: 12)),
                        ),
                    ],
                  ],
                ],
              ),
            ),
            actions: [
              TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('Cancelar'),
              ),
              FilledButton.icon(
                onPressed: selected.isEmpty || invalidQuantity
                    ? null
                    : () => Navigator.pop(context, {
                          'items': [
                            for (final line in selected)
                              {
                                'productId': line.productId,
                                'quantity': _number(line.quantity.text),
                              }
                          ],
                          'reason': _reason.text.trim(),
                          'refundMethod': refundDue > 0 ? _refundMethod : '',
                        }),
                icon: const Icon(Icons.assignment_return_outlined),
                label: const Text('Registrar devolución'),
              ),
            ],
          );
        },
      );

  Widget _returnPreviewRow(String label, String value, {bool strong = false}) =>
      Padding(
        padding: const EdgeInsets.symmetric(vertical: 3),
        child: Row(children: [
          Expanded(
              child: Text(label,
                  style: TextStyle(
                      fontWeight:
                          strong ? FontWeight.w800 : FontWeight.normal))),
          Text(value,
              style: TextStyle(
                  fontWeight: strong ? FontWeight.w800 : FontWeight.normal)),
        ]),
      );
}

class _SaleReturnLine {
  _SaleReturnLine({
    required this.productId,
    required this.name,
    required this.unitPrice,
    required this.soldQuantity,
    required this.returnedQuantity,
  }) : quantity = TextEditingController(text: '0');
  final String productId, name;
  final double unitPrice, soldQuantity, returnedQuantity;
  final TextEditingController quantity;
  double get availableQuantity => soldQuantity - returnedQuantity;
}

class _TransactionDialog extends StatefulWidget {
  const _TransactionDialog(
      {required this.products, required this.contacts, required this.purchase});
  final List<Map<String, dynamic>> products;
  final List<Map<String, dynamic>> contacts;
  final bool purchase;
  @override
  State<_TransactionDialog> createState() => _TransactionDialogState();
}

class _TransactionDialogState extends State<_TransactionDialog> {
  late final List<_TransactionLine> _lines = [
    _TransactionLine(widget.products.first['id'] as String)
  ];
  final _party = TextEditingController();
  final _note = TextEditingController();
  final _paid = TextEditingController(text: '0');
  final _reminderPhone = TextEditingController();
  bool _isCredit = false;
  String _paymentMethod = 'cash';
  bool _autoReminders = true;
  int _installmentCount = 1;
  String _installmentFrequency = 'monthly';
  DateTime _firstDueDate = DateTime.now().add(const Duration(days: 30));
  String? _contactId;

  @override
  void dispose() {
    for (final line in _lines) {
      line.dispose();
    }
    _party.dispose();
    _note.dispose();
    _paid.dispose();
    _reminderPhone.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) =>
      StatefulBuilder(builder: (context, refresh) {
        final priceField = widget.purchase ? 'cost' : 'price';
        final items = <Map<String, dynamic>>[];
        var total = 0.0;
        var stockValid = true;
        final requestedById = <String, double>{};
        for (final line in _lines) {
          final product =
              widget.products.firstWhere((p) => p['id'] == line.productId);
          final quantity = _number(line.quantity.text);
          requestedById.update(line.productId, (n) => n + quantity,
              ifAbsent: () => quantity);
          final unit = _number(product[priceField]);
          total += unit * quantity;
          items.add({
            'productId': line.productId,
            'name': product['name'],
            'quantity': quantity,
            'unitPrice': unit,
          });
        }
        if (!widget.purchase) {
          for (final entry in requestedById.entries) {
            final product =
                widget.products.firstWhere((p) => p['id'] == entry.key);
            if (_number(product['stock']) < entry.value) stockValid = false;
          }
        }
        final paid = widget.purchase
            ? total
            : _isCredit
                ? _number(_paid.text)
                : total;
        final outstanding = (total - paid).clamp(0, total).toDouble();
        final pocketSurcharge =
            !widget.purchase && !_isCredit && _paymentMethod == 'pocket'
                ? double.parse((total * .05).toStringAsFixed(2))
                : 0.0;
        final pocketAmount = total + pocketSurcharge;
        final previewSchedule = _buildPaymentSchedule(
          balance: outstanding,
          count: _installmentCount,
          frequency: _installmentFrequency,
          firstDueDate: _firstDueDate,
        );
        final phoneDigits = _reminderPhone.text.replaceAll(RegExp(r'\D'), '');
        final reminderPhoneValid =
            phoneDigits.length >= 7 && phoneDigits.length <= 15;
        final screenHeight = MediaQuery.sizeOf(context).height;
        return AlertDialog(
          title: Text(widget.purchase ? 'Registrar compra' : 'Registrar venta'),
          content: SizedBox(
            width: math.min(520.0, MediaQuery.sizeOf(context).width - 40),
            height: screenHeight * .68,
            child: ListView(children: [
              for (var i = 0; i < _lines.length; i++)
                Card(
                  elevation: 0,
                  color: const Color(0xfff4f7fc),
                  child: Padding(
                    padding: const EdgeInsets.all(12),
                    child: Column(children: [
                      Row(children: [
                        Expanded(
                          child: DropdownButtonFormField<String>(
                            initialValue: _lines[i].productId,
                            isExpanded: true,
                            decoration: InputDecoration(
                                labelText: 'Producto ${i + 1}', isDense: true),
                            items: widget.products
                                .map((p) => DropdownMenuItem(
                                    value: p['id'] as String,
                                    child: Text(
                                        '${p['name']} · stock ${_number(p['stock'])}',
                                        overflow: TextOverflow.ellipsis)))
                                .toList(),
                            onChanged: (value) {
                              if (value == null) return;
                              refresh(() => _lines[i].productId = value);
                            },
                          ),
                        ),
                        if (_lines.length > 1)
                          IconButton(
                            tooltip: 'Quitar producto',
                            onPressed: () => refresh(() {
                              _lines.removeAt(i).dispose();
                            }),
                            icon: const Icon(Icons.delete_outline),
                          ),
                      ]),
                      const SizedBox(height: 8),
                      TextField(
                        controller: _lines[i].quantity,
                        keyboardType: const TextInputType.numberWithOptions(
                            decimal: true),
                        decoration: const InputDecoration(
                            labelText: 'Cantidad', isDense: true),
                        onChanged: (_) => refresh(() {}),
                      ),
                    ]),
                  ),
                ),
              Align(
                alignment: Alignment.centerLeft,
                child: TextButton.icon(
                  onPressed: () => refresh(() => _lines.add(
                      _TransactionLine(widget.products.first['id'] as String))),
                  icon: const Icon(Icons.add),
                  label: const Text('Agregar otro producto'),
                ),
              ),
              _field(_party, widget.purchase ? 'Proveedor' : 'Cliente',
                  required: !widget.purchase && _isCredit),
              _field(_note, 'Nota / descripción'),
              if (!widget.purchase && !_isCredit) ...[
                DropdownButtonFormField<String>(
                  initialValue: _paymentMethod,
                  decoration:
                      const InputDecoration(labelText: 'Método de pago'),
                  items: const [
                    DropdownMenuItem(value: 'cash', child: Text('Efectivo')),
                    DropdownMenuItem(
                        value: 'pocket', child: Text('Link de pago · Pocket')),
                  ],
                  onChanged: (value) => refresh(() {
                    if (value != null) _paymentMethod = value;
                  }),
                ),
                if (_paymentMethod == 'pocket') ...[
                  const SizedBox(height: 8),
                  TextField(
                    controller: _reminderPhone,
                    keyboardType: TextInputType.phone,
                    decoration: const InputDecoration(
                      labelText: 'WhatsApp del cliente',
                      hintText: 'Incluye el código de país',
                      prefixIcon: Icon(Icons.phone_outlined),
                      isDense: true,
                    ),
                    onChanged: (_) => refresh(() {}),
                  ),
                  if (!reminderPhoneValid)
                    const Padding(
                      padding: EdgeInsets.only(top: 5),
                      child: Align(
                        alignment: Alignment.centerLeft,
                        child: Text(
                          'Escribe un número con código de país para enviar el enlace.',
                          style: TextStyle(color: Colors.red, fontSize: 12),
                        ),
                      ),
                    ),
                ],
              ],
              if (!widget.purchase)
                CheckboxListTile(
                  contentPadding: EdgeInsets.zero,
                  value: _isCredit,
                  title: const Text('Venta al crédito'),
                  subtitle: const Text(
                      'Define cuotas, vencimientos y avisos de cobro.'),
                  onChanged: (value) => refresh(() {
                    _isCredit = value ?? false;
                    if (_isCredit &&
                        widget.contacts.isNotEmpty &&
                        _contactId == null) {
                      _contactId = widget.contacts.first['id']?.toString();
                      _party.text =
                          widget.contacts.first['name']?.toString() ?? '';
                      _reminderPhone.text =
                          widget.contacts.first['phone']?.toString() ?? '';
                    }
                  }),
                ),
              if (_isCredit && widget.contacts.isNotEmpty) ...[
                DropdownButtonFormField<String>(
                  initialValue: _contactId,
                  decoration:
                      const InputDecoration(labelText: 'Cliente registrado'),
                  items: widget.contacts
                      .where((contact) => contact['id'] != null)
                      .map((contact) => DropdownMenuItem<String>(
                            value: contact['id'].toString(),
                            child:
                                Text(contact['name']?.toString() ?? 'Cliente'),
                          ))
                      .toList(),
                  onChanged: (value) => refresh(() {
                    _contactId = value;
                    if (value == null) return;
                    final contact = widget.contacts
                        .firstWhere((entry) => entry['id'].toString() == value);
                    _party.text = contact['name']?.toString() ?? '';
                    _reminderPhone.text = contact['phone']?.toString() ?? '';
                  }),
                ),
                const SizedBox(height: 8),
              ],
              if (_isCredit) ...[
                _field(_paid, 'Abono recibido hoy', number: true),
                if (paid < 0 || paid > total)
                  const Align(
                    alignment: Alignment.centerLeft,
                    child: Text('El abono debe estar entre cero y el total.',
                        style: TextStyle(color: Colors.red)),
                  ),
                if (outstanding > 0) ...[
                  const SizedBox(height: 4),
                  Card(
                    elevation: 0,
                    color: const Color(0xfff4f7fc),
                    child: Padding(
                      padding: const EdgeInsets.all(14),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          const Text('Plan de cuotas',
                              style: TextStyle(fontWeight: FontWeight.w800)),
                          const SizedBox(height: 10),
                          Wrap(spacing: 10, runSpacing: 8, children: [
                            SizedBox(
                              width: 140,
                              child: DropdownButtonFormField<int>(
                                initialValue: _installmentCount,
                                decoration: const InputDecoration(
                                    labelText: 'Número de cuotas',
                                    isDense: true),
                                items: List.generate(12, (index) => index + 1)
                                    .map((count) => DropdownMenuItem(
                                        value: count,
                                        child: Text(
                                            '$count ${count == 1 ? 'cuota' : 'cuotas'}')))
                                    .toList(),
                                onChanged: (value) => refresh(() {
                                  if (value != null) _installmentCount = value;
                                }),
                              ),
                            ),
                            SizedBox(
                              width: 150,
                              child: DropdownButtonFormField<String>(
                                initialValue: _installmentFrequency,
                                decoration: const InputDecoration(
                                    labelText: 'Frecuencia', isDense: true),
                                items: const [
                                  DropdownMenuItem(
                                      value: 'weekly', child: Text('Semanal')),
                                  DropdownMenuItem(
                                      value: 'biweekly',
                                      child: Text('Cada 2 semanas')),
                                  DropdownMenuItem(
                                      value: 'monthly', child: Text('Mensual')),
                                ],
                                onChanged: (value) => refresh(() {
                                  if (value != null)
                                    _installmentFrequency = value;
                                }),
                              ),
                            ),
                            OutlinedButton.icon(
                              onPressed: () async {
                                final selected = await showDatePicker(
                                  context: context,
                                  initialDate: _firstDueDate,
                                  firstDate: DateTime.now()
                                      .add(const Duration(days: 1)),
                                  lastDate: DateTime.now()
                                      .add(const Duration(days: 3650)),
                                  helpText: 'Selecciona el primer vencimiento',
                                );
                                if (selected != null) {
                                  refresh(() => _firstDueDate = DateTime(
                                      selected.year,
                                      selected.month,
                                      selected.day,
                                      9));
                                }
                              },
                              icon: const Icon(Icons.calendar_month_outlined),
                              label: Text(
                                  'Primera: ${DateFormat('dd/MM/yyyy').format(_firstDueDate)}'),
                            ),
                          ]),
                          const SizedBox(height: 8),
                          TextField(
                            controller: _reminderPhone,
                            keyboardType: TextInputType.phone,
                            decoration: const InputDecoration(
                                labelText: 'WhatsApp del cliente',
                                hintText: 'Incluye el código de país',
                                prefixIcon: Icon(Icons.phone_outlined),
                                isDense: true),
                            onChanged: (_) => refresh(() {}),
                          ),
                          if (_autoReminders && !reminderPhoneValid)
                            const Padding(
                              padding: EdgeInsets.only(top: 5),
                              child: Text(
                                  'Añade un teléfono con código de país para programar los avisos.',
                                  style: TextStyle(
                                      color: Colors.red, fontSize: 12)),
                            ),
                          SwitchListTile.adaptive(
                            contentPadding: EdgeInsets.zero,
                            value: _autoReminders,
                            title: const Text('Aviso automático por WhatsApp'),
                            subtitle: const Text(
                                'Se enviará 1 día antes de cada vencimiento.'),
                            onChanged: (value) =>
                                refresh(() => _autoReminders = value),
                          ),
                          Text(
                              'Las cuotas se distribuyen sobre el saldo pendiente. El servidor debe estar activo y el WhatsApp de esta cuenta conectado al momento del envío.',
                              style: TextStyle(
                                  color: Colors.black54, fontSize: 12)),
                          if (previewSchedule.isNotEmpty) ...[
                            const Divider(height: 18),
                            for (final installment in previewSchedule.take(4))
                              Padding(
                                padding:
                                    const EdgeInsets.symmetric(vertical: 3),
                                child: Row(children: [
                                  Expanded(
                                      child: Text(
                                          'Cuota ${installment['number']}')),
                                  Text(_money(_number(installment['amount']))),
                                  const SizedBox(width: 12),
                                  Text(DateFormat('dd/MM/yyyy').format(
                                      DateTime.parse(
                                              installment['dueAt'] as String)
                                          .toLocal())),
                                ]),
                              ),
                            if (previewSchedule.length > 4)
                              Text('Y ${previewSchedule.length - 4} cuotas más',
                                  style: const TextStyle(
                                      color: Colors.black54, fontSize: 12)),
                          ],
                        ],
                      ),
                    ),
                  ),
                ],
              ],
              if (!stockValid)
                const Align(
                  alignment: Alignment.centerLeft,
                  child: Text(
                      'La cantidad supera las existencias de uno de los productos.',
                      style: TextStyle(color: Colors.red)),
                ),
              const SizedBox(height: 10),
              Align(
                alignment: Alignment.centerRight,
                child: Column(
                    crossAxisAlignment: CrossAxisAlignment.end,
                    children: [
                      if (pocketSurcharge > 0)
                        Text('Subtotal: ${_money(total)}',
                            style: const TextStyle(color: Colors.black54)),
                      if (pocketSurcharge > 0)
                        Text(
                            'Recargo de pago Pocket (5%): ${_money(pocketSurcharge)}',
                            style: const TextStyle(color: Colors.black54)),
                      Text('Total: ${_money(pocketAmount)}',
                          style: const TextStyle(
                              fontWeight: FontWeight.w800, fontSize: 18)),
                      if (!widget.purchase && _isCredit)
                        Text('Pendiente: ${_money(total - paid)}',
                            style: const TextStyle(
                                color: _brand, fontWeight: FontWeight.w700)),
                    ]),
              ),
            ]),
          ),
          actions: [
            TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('Cancelar')),
            FilledButton(
              onPressed: () {
                if (items.isEmpty ||
                    items.any((item) => _number(item['quantity']) <= 0)) return;
                if (!stockValid || (_isCredit && (paid < 0 || paid > total)))
                  return;
                if (_isCredit && _party.text.trim().isEmpty) return;
                if (!widget.purchase &&
                    !_isCredit &&
                    _paymentMethod == 'pocket' &&
                    !reminderPhoneValid) return;
                if (_isCredit &&
                    outstanding > 0 &&
                    _autoReminders &&
                    !reminderPhoneValid) return;
                final names = items.map((item) => item['name']).join(', ');
                Navigator.pop(context, {
                  'id': DateTime.now().microsecondsSinceEpoch.toString(),
                  'label': '${widget.purchase ? 'Compra' : 'Venta'} · $names',
                  'number':
                      '${widget.purchase ? 'C' : 'V'}-${DateTime.now().millisecondsSinceEpoch.toString().substring(7)}',
                  'items': items,
                  'subtotal': total,
                  'pocketSurcharge': pocketSurcharge,
                  'pocketAmount': pocketAmount,
                  'isCredit': !widget.purchase && _isCredit,
                  'paymentMethod': widget.purchase
                      ? 'cash'
                      : _isCredit
                          ? 'credit'
                          : _paymentMethod,
                  if (!widget.purchase && _isCredit && _contactId != null)
                    'customerId': _contactId,
                  'amountPaid': paid,
                  'installmentCount': _installmentCount,
                  'installmentFrequency': _installmentFrequency,
                  'firstDueAt': DateTime(_firstDueDate.year,
                          _firstDueDate.month, _firstDueDate.day, 9)
                      .toIso8601String(),
                  'autoReminders': !widget.purchase &&
                      _isCredit &&
                      outstanding > 0 &&
                      _autoReminders,
                  'reminderPhone': phoneDigits,
                  'party': _party.text.trim(),
                  'detail': _note.text.trim(),
                  'date': DateFormat('dd/MM/yyyy HH:mm').format(DateTime.now()),
                });
              },
              child: const Text('Registrar'),
            ),
          ],
        );
      });
}

class _TransactionLine {
  _TransactionLine(this.productId)
      : quantity = TextEditingController(text: '1');
  String productId;
  final TextEditingController quantity;
  void dispose() => quantity.dispose();
}

class _ReceiptCard extends StatelessWidget {
  const _ReceiptCard({
    required this.sale,
    required this.businessName,
    required this.logo,
    required this.taxId,
    required this.address,
    required this.phone,
    required this.email,
  });
  final Map<String, dynamic> sale;
  final String businessName, logo, taxId, address, phone, email;

  String _amount(double value) => NumberFormat.currency(
        locale: _currencyLocales[sale['currency']?.toString()] ?? 'es_NI',
        symbol: _currencySymbols[sale['currency']?.toString()] ?? 'C\$',
        decimalDigits: 2,
      ).format(value);

  @override
  Widget build(BuildContext context) {
    final items = sale['items'] is List
        ? (sale['items'] as List).whereType<Map>().toList()
        : const <Map>[];
    final total = _number(sale['total']);
    final surcharge = _number(sale['pocketSurcharge']);
    final subtotal = sale['subtotal'] == null
        ? total - surcharge
        : _number(sale['subtotal']);
    final paid = _number(sale['amountPaid']);
    final balance = _number(sale['balance']);
    final returnedAmount = _number(sale['returnedAmount']);
    final logoBytes = logo.startsWith('data:image/') && logo.contains(',')
        ? base64Decode(logo.split(',').last)
        : null;
    return Material(
      color: Colors.white,
      child: Container(
        width: 410,
        padding: const EdgeInsets.all(22),
        decoration: BoxDecoration(
          color: Colors.white,
          border: Border.all(color: const Color(0xffe3e9f0)),
          borderRadius: BorderRadius.circular(16),
        ),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          mainAxisSize: MainAxisSize.min,
          children: [
            if (logoBytes != null)
              Center(
                child: Padding(
                  padding: const EdgeInsets.only(bottom: 10),
                  child: ClipRRect(
                    borderRadius: BorderRadius.circular(12),
                    child: Image.memory(logoBytes,
                        width: 76, height: 76, fit: BoxFit.cover),
                  ),
                ),
              ),
            Text(businessName,
                textAlign: TextAlign.center,
                style: const TextStyle(
                    fontSize: 20, fontWeight: FontWeight.w800, color: _brand)),
            if (taxId.isNotEmpty) _receiptCentered('RUC: $taxId'),
            if (address.isNotEmpty) _receiptCentered(address),
            if (phone.isNotEmpty) _receiptCentered(phone),
            if (email.isNotEmpty) _receiptCentered(email),
            const SizedBox(height: 12),
            const Divider(height: 1),
            const SizedBox(height: 10),
            const Text('COMPROBANTE DE VENTA',
                textAlign: TextAlign.center,
                style:
                    TextStyle(fontWeight: FontWeight.w800, letterSpacing: .6)),
            const SizedBox(height: 8),
            _receiptRow('N.º',
                sale['number']?.toString() ?? sale['id']?.toString() ?? ''),
            _receiptRow('Fecha', sale['date']?.toString() ?? ''),
            _receiptRow(
                'Cliente',
                sale['party']?.toString().isNotEmpty == true
                    ? sale['party'].toString()
                    : 'Cliente'),
            const SizedBox(height: 10),
            const Divider(height: 1),
            const SizedBox(height: 8),
            for (final item in items) ...[
              Row(crossAxisAlignment: CrossAxisAlignment.start, children: [
                Expanded(
                  child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(item['name']?.toString() ?? 'Producto',
                            style:
                                const TextStyle(fontWeight: FontWeight.w600)),
                        Text(
                            '${_number(item['quantity'])} × ${_amount(_number(item['unitPrice']))}',
                            style: const TextStyle(
                                color: Colors.black54, fontSize: 12)),
                      ]),
                ),
                const SizedBox(width: 12),
                Text(_amount(_number(item['total'])),
                    style: const TextStyle(fontWeight: FontWeight.w600)),
              ]),
              const SizedBox(height: 8),
            ],
            const Divider(height: 1),
            const SizedBox(height: 8),
            _receiptRow('Subtotal', _amount(subtotal)),
            if (surcharge > 0)
              _receiptRow('Recargo de pago Pocket (5%)', _amount(surcharge)),
            if (returnedAmount > 0)
              _receiptRow(
                  'Devoluciones aplicadas', '-${_amount(returnedAmount)}'),
            const SizedBox(height: 4),
            _receiptRow('TOTAL', _amount(total), strong: true),
            const SizedBox(height: 8),
            _receiptRow(
                'Forma de pago',
                switch (sale['paymentMethod']) {
                  'pocket' => 'Poket · confirmado',
                  'credit' => 'Crédito',
                  'cash' => 'Efectivo',
                  _ => sale['paymentMethod']?.toString() ?? 'Efectivo',
                }),
            if (balance > 0) ...[
              _receiptRow('Abonado', _amount(paid)),
              _receiptRow('Saldo pendiente', _amount(balance), strong: true),
            ],
            if (sale['detail']?.toString().trim().isNotEmpty == true) ...[
              const SizedBox(height: 8),
              Text(sale['detail'].toString(),
                  style: const TextStyle(color: Colors.black54, fontSize: 12)),
            ],
            const SizedBox(height: 14),
            const Text('Gracias por tu compra',
                textAlign: TextAlign.center,
                style: TextStyle(
                    color: Colors.black54, fontWeight: FontWeight.w600)),
          ],
        ),
      ),
    );
  }

  Widget _receiptCentered(String value) => Padding(
        padding: const EdgeInsets.only(top: 3),
        child: Text(value,
            textAlign: TextAlign.center,
            style: const TextStyle(fontSize: 12, color: Colors.black54)),
      );

  Widget _receiptRow(String label, String value, {bool strong = false}) =>
      Padding(
        padding: const EdgeInsets.symmetric(vertical: 3),
        child: Row(children: [
          Expanded(
              child: Text(label,
                  style: TextStyle(
                      fontSize: 12,
                      fontWeight:
                          strong ? FontWeight.w800 : FontWeight.normal))),
          const SizedBox(width: 10),
          Flexible(
              child: Text(value,
                  textAlign: TextAlign.right,
                  style: TextStyle(
                      fontSize: 12,
                      fontWeight:
                          strong ? FontWeight.w800 : FontWeight.normal))),
        ]),
      );
}

class _PurchaseReceiptCard extends StatelessWidget {
  const _PurchaseReceiptCard({
    required this.purchase,
    required this.businessName,
    required this.logo,
    required this.taxId,
    required this.address,
    required this.phone,
    required this.email,
  });

  final Map<String, dynamic> purchase;
  final String businessName, logo, taxId, address, phone, email;

  String _amount(double value) => NumberFormat.currency(
        locale: _currencyLocales[purchase['currency']?.toString()] ?? 'es_NI',
        symbol: _currencySymbols[purchase['currency']?.toString()] ?? 'C\$',
        decimalDigits: 2,
      ).format(value);

  @override
  Widget build(BuildContext context) {
    final items = purchase['items'] is List
        ? (purchase['items'] as List).whereType<Map>().toList()
        : const <Map>[];
    final total = _number(purchase['total']);
    final subtotal =
        purchase['subtotal'] == null ? total : _number(purchase['subtotal']);
    final logoBytes = logo.startsWith('data:image/') && logo.contains(',')
        ? base64Decode(logo.split(',').last)
        : null;

    return Material(
      color: Colors.white,
      child: Container(
        width: 350,
        padding: const EdgeInsets.fromLTRB(20, 22, 20, 18),
        decoration: BoxDecoration(
          color: Colors.white,
          border: Border.all(color: const Color(0xffe3e9f0)),
          borderRadius: BorderRadius.circular(12),
        ),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            if (logoBytes != null)
              Center(
                child: Padding(
                  padding: const EdgeInsets.only(bottom: 9),
                  child: ClipRRect(
                    borderRadius: BorderRadius.circular(8),
                    child: Image.memory(logoBytes,
                        width: 58, height: 58, fit: BoxFit.cover),
                  ),
                ),
              ),
            Text(
              businessName.toUpperCase(),
              textAlign: TextAlign.center,
              style: const TextStyle(
                  fontSize: 18,
                  fontWeight: FontWeight.w900,
                  letterSpacing: .5,
                  color: _brand),
            ),
            if (taxId.isNotEmpty) _centered('RUC: $taxId'),
            if (address.isNotEmpty) _centered(address),
            if (phone.isNotEmpty) _centered(phone),
            if (email.isNotEmpty) _centered(email),
            const SizedBox(height: 10),
            _separator(),
            const SizedBox(height: 8),
            const Text('COMPROBANTE DE COMPRA',
                textAlign: TextAlign.center,
                style: TextStyle(fontWeight: FontWeight.w900, fontSize: 14)),
            const SizedBox(height: 3),
            const Text('REGISTRO DE MERCANCÍA',
                textAlign: TextAlign.center,
                style: TextStyle(
                    fontSize: 10,
                    letterSpacing: 1.1,
                    color: Colors.black54,
                    fontWeight: FontWeight.w700)),
            const SizedBox(height: 8),
            _row(
                'N.º',
                purchase['number']?.toString() ??
                    purchase['id']?.toString() ??
                    ''),
            _row('Fecha', purchase['date']?.toString() ?? ''),
            _row(
                'Proveedor',
                purchase['party']?.toString().trim().isNotEmpty == true
                    ? purchase['party'].toString()
                    : 'Proveedor'),
            const SizedBox(height: 7),
            _separator(),
            const SizedBox(height: 8),
            for (final item in items) ...[
              Row(crossAxisAlignment: CrossAxisAlignment.start, children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(item['name']?.toString() ?? 'Producto',
                          style: const TextStyle(
                              fontWeight: FontWeight.w700, fontSize: 12)),
                      Text(
                        '${_number(item['quantity'])} × ${_amount(_number(item['unitPrice']))}',
                        style: const TextStyle(
                            color: Colors.black54, fontSize: 11),
                      ),
                    ],
                  ),
                ),
                const SizedBox(width: 8),
                Text(_amount(_number(item['total'])),
                    style: const TextStyle(
                        fontSize: 12, fontWeight: FontWeight.w700)),
              ]),
              const SizedBox(height: 7),
            ],
            _separator(),
            const SizedBox(height: 7),
            _row('Unidades recibidas',
                _number(purchase['quantity']).toStringAsFixed(0)),
            _row('Subtotal', _amount(subtotal)),
            const SizedBox(height: 3),
            _row('TOTAL COMPRA', _amount(total), strong: true),
            const SizedBox(height: 7),
            _separator(),
            const SizedBox(height: 8),
            _row(
                'Forma de pago',
                switch (purchase['paymentMethod']) {
                  'pocket' => 'Poket',
                  'card' => 'Tarjeta',
                  'transfer' => 'Transferencia',
                  'cash' => 'Efectivo',
                  _ => purchase['paymentMethod']?.toString() ?? 'Efectivo',
                }),
            if (purchase['detail']?.toString().trim().isNotEmpty == true) ...[
              const SizedBox(height: 7),
              Text('Nota: ${purchase['detail']}',
                  style: const TextStyle(fontSize: 11, color: Colors.black54)),
            ],
            const SizedBox(height: 13),
            const Text('Mercadería recibida y registrada',
                textAlign: TextAlign.center,
                style: TextStyle(fontWeight: FontWeight.w700, fontSize: 11)),
            const SizedBox(height: 3),
            const Text('Gracias por tu trabajo',
                textAlign: TextAlign.center,
                style: TextStyle(fontSize: 10, color: Colors.black54)),
          ],
        ),
      ),
    );
  }

  Widget _centered(String value) => Padding(
        padding: const EdgeInsets.only(top: 3),
        child: Text(value,
            textAlign: TextAlign.center,
            style: const TextStyle(fontSize: 10, color: Colors.black54)),
      );

  Widget _separator() => const Text(
        '------------------------------------------',
        maxLines: 1,
        overflow: TextOverflow.clip,
        textAlign: TextAlign.center,
        style: TextStyle(
            fontSize: 9,
            height: 1,
            letterSpacing: .2,
            color: Color(0xff8a929c)),
      );

  Widget _row(String label, String value, {bool strong = false}) => Padding(
        padding: const EdgeInsets.symmetric(vertical: 2),
        child: Row(children: [
          Expanded(
            child: Text(label,
                style: TextStyle(
                    fontSize: strong ? 13 : 11,
                    fontWeight: strong ? FontWeight.w900 : FontWeight.normal)),
          ),
          const SizedBox(width: 8),
          Flexible(
            child: Text(value,
                textAlign: TextAlign.right,
                style: TextStyle(
                    fontSize: strong ? 13 : 11,
                    fontWeight: strong ? FontWeight.w900 : FontWeight.normal)),
          ),
        ]),
      );
}

class _ExpenseDialog extends StatefulWidget {
  const _ExpenseDialog();
  @override
  State<_ExpenseDialog> createState() => _ExpenseDialogState();
}

class _ExpenseDialogState extends State<_ExpenseDialog> {
  final _description = TextEditingController();
  final _amount = TextEditingController();
  final _category = TextEditingController();
  @override
  Widget build(BuildContext context) => AlertDialog(
          title: const Text('Nuevo gasto'),
          content: SizedBox(
              width: 430,
              child: Column(mainAxisSize: MainAxisSize.min, children: [
                _field(_description, 'Descripción', required: true),
                _field(_category, 'Categoría'),
                _field(_amount, 'Monto', number: true, required: true)
              ])),
          actions: [
            TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('Cancelar')),
            FilledButton(
                onPressed: () {
                  if (_description.text.trim().isEmpty ||
                      _number(_amount.text) <= 0) return;
                  Navigator.pop(context, {
                    'label': _description.text.trim(),
                    'number':
                        'G-${DateTime.now().millisecondsSinceEpoch.toString().substring(7)}',
                    'detail': _category.text.trim(),
                    'total': _number(_amount.text),
                    'date':
                        DateFormat('dd/MM/yyyy HH:mm').format(DateTime.now())
                  });
                },
                child: const Text('Guardar gasto'))
          ]);
}

class _MetricCard extends StatelessWidget {
  const _MetricCard(this.label, this.value, this.icon, this.tint);
  final String label, value;
  final IconData icon;
  final Color tint;
  @override
  Widget build(BuildContext context) => Card(
      elevation: 0,
      child: Padding(
          padding: const EdgeInsets.all(16),
          child: Row(children: [
            CircleAvatar(
                backgroundColor: tint, child: Icon(icon, color: _brand)),
            const SizedBox(width: 13),
            Expanded(
                child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    mainAxisAlignment: MainAxisAlignment.center,
                    children: [
                  Text(label,
                      style:
                          const TextStyle(color: Colors.black54, fontSize: 12)),
                  const SizedBox(height: 5),
                  Text(value,
                      overflow: TextOverflow.ellipsis,
                      style: const TextStyle(
                          fontSize: 19,
                          fontWeight: FontWeight.w800,
                          color: _ink))
                ]))
          ])));
}

class _Panel extends StatelessWidget {
  const _Panel({required this.title, required this.child, this.trailing});
  final String title;
  final Widget child;
  final Widget? trailing;
  @override
  Widget build(BuildContext context) => Card(
      elevation: 0,
      child: Padding(
          padding: const EdgeInsets.all(18),
          child:
              Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
            Row(children: [
              Expanded(
                  child: Text(title,
                      style: const TextStyle(
                          fontSize: 17, fontWeight: FontWeight.w800))),
              if (trailing != null) trailing!
            ]),
            const SizedBox(height: 12),
            child
          ])));
}

class _QuickAction extends StatelessWidget {
  const _QuickAction(
      {required this.icon, required this.label, required this.onTap});
  final IconData icon;
  final String label;
  final VoidCallback onTap;
  @override
  Widget build(BuildContext context) => ListTile(
      contentPadding: EdgeInsets.zero,
      leading: Icon(icon, color: _brand),
      title: Text(label, style: const TextStyle(fontWeight: FontWeight.w600)),
      trailing: const Icon(Icons.chevron_right),
      onTap: onTap);
}

class _PageHeading extends StatelessWidget {
  const _PageHeading(this.title, this.subtitle);
  final String title, subtitle;
  @override
  Widget build(BuildContext context) =>
      Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
        Text(title,
            style: const TextStyle(
                fontSize: 28, fontWeight: FontWeight.w800, color: _ink)),
        const SizedBox(height: 4),
        Text(subtitle, style: const TextStyle(color: Colors.black54))
      ]);
}

class _Empty extends StatelessWidget {
  const _Empty({required this.icon, required this.text});
  final IconData icon;
  final String text;
  @override
  Widget build(BuildContext context) => Card(
      elevation: 0,
      child: Padding(
          padding: const EdgeInsets.all(32),
          child: Center(
              child: Column(children: [
            Icon(icon, size: 36, color: _brand),
            const SizedBox(height: 12),
            Text(text,
                textAlign: TextAlign.center,
                style: const TextStyle(color: Colors.black54))
          ]))));
}

class _BrandMark extends StatelessWidget {
  const _BrandMark();
  @override
  Widget build(BuildContext context) =>
      Row(mainAxisSize: MainAxisSize.min, children: [
        Container(
            width: 40,
            height: 40,
            decoration: BoxDecoration(
                color: _brand, borderRadius: BorderRadius.circular(13)),
            child: ClipRRect(
                borderRadius: BorderRadius.circular(13),
                child: Image.asset('assets/branding/sync_connect_logo.png',
                    width: 40, height: 40, fit: BoxFit.cover))),
        const SizedBox(width: 10),
        const Text('Sync Connect',
            style: TextStyle(
                fontSize: 18, fontWeight: FontWeight.w800, color: _ink))
      ]);
}

class _Loading extends StatelessWidget {
  const _Loading({this.message = 'Cargando…'});
  final String message;
  @override
  Widget build(BuildContext context) => Scaffold(
        body: Center(
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            const SizedBox(
                width: 32,
                height: 32,
                child: CircularProgressIndicator(strokeWidth: 3)),
            const SizedBox(height: 14),
            Text(message, style: const TextStyle(color: _ink)),
          ]),
        ),
      );
}

Widget _field(TextEditingController controller, String label,
        {bool number = false, bool required = false, int lines = 1}) =>
    Padding(
        padding: const EdgeInsets.only(bottom: 10),
        child: TextFormField(
          controller: controller,
          maxLines: lines,
          keyboardType: number
              ? const TextInputType.numberWithOptions(decimal: true)
              : TextInputType.text,
          decoration: InputDecoration(labelText: label),
          validator: required
              ? (v) => (v == null || v.trim().isEmpty)
                  ? 'Campo requerido'
                  : number && _number(v) < 0
                      ? 'Usa un valor válido'
                      : null
              : null,
        ));
double _number(dynamic value) =>
    value is num ? value.toDouble() : double.tryParse('$value') ?? 0;
