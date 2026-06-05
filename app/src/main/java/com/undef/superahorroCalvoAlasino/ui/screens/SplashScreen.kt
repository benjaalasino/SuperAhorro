package com.undef.superahorroCalvoAlasino.ui.screens

import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.collectAsState
import androidx.compose.runtime.getValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.navigation.NavController
import com.undef.superahorroCalvoAlasino.navigation.NavRoutes
import com.undef.superahorroCalvoAlasino.viewmodel.UsuarioViewModel

@Composable
fun SplashScreen(navController: NavController, usuarioViewModel: UsuarioViewModel) {
    // Observa el StateFlow reactivo (requerimiento de la tarea)
    val isLoggedIn by usuarioViewModel.isLoggedIn.collectAsState()

    // Lee el valor real desde DataStore una sola vez para decidir la navegación.
    // No usamos `isLoggedIn` directamente aquí porque su initialValue es false,
    // lo que causaría navegar a Login antes de que DataStore responda.
    LaunchedEffect(Unit) {
        val loggedIn = usuarioViewModel.checkInitialLoginState()
        navController.navigate(
            if (loggedIn) NavRoutes.Home.route else NavRoutes.Login.route
        ) {
            popUpTo(NavRoutes.Splash.route) { inclusive = true }
        }
    }

    Box(
        modifier = Modifier.fillMaxSize(),
        contentAlignment = Alignment.Center
    ) {
        CircularProgressIndicator(color = Color(0xFF2E7D32))
    }
}
