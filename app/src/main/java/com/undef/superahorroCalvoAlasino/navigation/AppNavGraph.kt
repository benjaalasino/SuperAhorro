package com.undef.superahorroCalvoAlasino.navigation

import android.content.Context
import androidx.compose.runtime.Composable
import androidx.compose.runtime.remember
import androidx.navigation.NavType
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.rememberNavController
import androidx.navigation.navArgument
import com.undef.superahorroCalvoAlasino.data.preferences.UserPreferencesRepository
import com.undef.superahorroCalvoAlasino.ui.screens.*
import com.undef.superahorroCalvoAlasino.viewmodel.CompraViewModel
import com.undef.superahorroCalvoAlasino.viewmodel.UsuarioViewModel

@Composable
fun AppNavGraph(context: Context? = null) {
    val navController = rememberNavController()
    val compraViewModel = remember { CompraViewModel(context) }
    val userPrefsRepo = remember { UserPreferencesRepository(context!!) }
    val usuarioViewModel = remember { UsuarioViewModel(context, userPrefsRepo) }

    NavHost(
        navController = navController,
        startDestination = NavRoutes.Splash.route
    ) {
        composable(NavRoutes.Splash.route) { SplashScreen(navController, usuarioViewModel) }
        composable(NavRoutes.Login.route) { LoginScreen(navController, usuarioViewModel, compraViewModel) }
        composable(NavRoutes.Registro.route) { RegistroScreen(navController, usuarioViewModel, compraViewModel) }
        composable(NavRoutes.Home.route) { HomeScreen(navController, compraViewModel) }
        composable(NavRoutes.Historial.route) { HistorialScreen(navController, compraViewModel) }
        composable(NavRoutes.Estadisticas.route) { EstadisticasScreen(navController, compraViewModel) }
        composable(NavRoutes.Perfil.route) { PerfilScreen(navController, usuarioViewModel, compraViewModel) }
        composable(NavRoutes.EditarPerfil.route) { EditarPerfilScreen(navController, usuarioViewModel) }
        composable(NavRoutes.Settings.route) { SettingsScreen(navController) }

        composable(NavRoutes.NuevaCompra.route) { NuevaCompraScreen(navController, compraViewModel) }

        composable(
            route = NavRoutes.DetalleCompra.route,
            arguments = listOf(navArgument("compraId") { type = NavType.IntType })
        ) { backStackEntry ->
            val compraId = backStackEntry.arguments?.getInt("compraId") ?: 0
            DetalleCompraScreen(navController, compraId, compraViewModel)
        }

        composable(
            route = NavRoutes.NuevoProducto.route,
            arguments = listOf(navArgument("compraId") { type = NavType.IntType })
        ) { backStackEntry ->
            val compraId = backStackEntry.arguments?.getInt("compraId") ?: 0
            NuevoProductoScreen(navController, compraId, compraViewModel)
        }
    }
}

