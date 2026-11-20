// Тесты обращаются к internal-типам библиотеки. Сборки в репозитории не подписываются,
// поэтому InternalsVisibleTo указан без PublicKey.
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("Aid.Xlsx.Tests")]