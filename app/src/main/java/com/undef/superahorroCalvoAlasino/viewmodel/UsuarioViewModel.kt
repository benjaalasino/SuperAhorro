package com.undef.superahorroCalvoAlasino.viewmodel

import android.content.Context
import android.content.SharedPreferences
import androidx.compose.runtime.mutableStateOf
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import com.undef.superahorroCalvoAlasino.data.preferences.UserPreferencesRepository
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.stateIn
import kotlinx.coroutines.launch

data class Usuario(
    val nombre: String = "",
    val email: String = "",
    val password: String = ""
)

class UsuarioViewModel(
    private var context: Context? = null,
    private val repo: UserPreferencesRepository
) : ViewModel() {

    private val _usuario = mutableStateOf(Usuario())
    val usuario = _usuario

    // Contraseña sigue en SharedPreferences solo para validación de login
    private var sharedPreferences: SharedPreferences? = null

    val isLoggedIn: StateFlow<Boolean> = repo.isLoggedIn
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5_000), false)

    init {
        if (context != null) {
            initializePreferences()
        }
    }

    fun setContext(context: Context) {
        this.context = context
        initializePreferences()
    }

    private fun initializePreferences() {
        if (sharedPreferences == null) {
            context?.let {
                sharedPreferences = it.getSharedPreferences("super_ahorro_prefs", Context.MODE_PRIVATE)
                cargarUsuarioGuardado()
            }
        }
    }

    fun cargarUsuarioGuardado() {
        if (sharedPreferences == null) {
            initializePreferences()
        }
        sharedPreferences?.let {
            val nombre = it.getString("nombre", "") ?: ""
            val email = it.getString("email", "") ?: ""
            val password = it.getString("password", "") ?: ""
            if (nombre.isNotEmpty() || email.isNotEmpty() || password.isNotEmpty()) {
                _usuario.value = Usuario(nombre = nombre, email = email, password = password)
            }
        }
    }

    fun registrarUsuario(nombre: String, email: String, password: String) {
        if (sharedPreferences == null) {
            initializePreferences()
        }
        _usuario.value = Usuario(nombre = nombre, email = email, password = password)
        guardarUsuarioEnPreferencias()
        viewModelScope.launch {
            repo.saveSession(email, nombre)
        }
    }

    fun actualizarPerfil(nombre: String, email: String, nuevaPassword: String) {
        _usuario.value = Usuario(nombre = nombre, email = email, password = nuevaPassword)
        guardarUsuarioEnPreferencias()
    }

    private fun guardarUsuarioEnPreferencias() {
        if (sharedPreferences == null) {
            initializePreferences()
        }
        sharedPreferences?.edit()?.apply {
            putString("nombre", _usuario.value.nombre)
            putString("email", _usuario.value.email)
            putString("password", _usuario.value.password)
            apply()
        }
    }

    fun obtenerUsuario(): Usuario {
        return _usuario.value
    }

    fun validarCredenciales(email: String, password: String): Boolean {
        if (_usuario.value.email.isEmpty()) {
            cargarUsuarioGuardado()
        }
        val usuarioGuardado = _usuario.value
        val result = usuarioGuardado.email == email && usuarioGuardado.password == password
        if (result) {
            viewModelScope.launch {
                repo.saveSession(email, usuarioGuardado.nombre)
            }
        }
        return result
    }

    fun cerrarSesion() {
        _usuario.value = Usuario()
        viewModelScope.launch {
            repo.clearSession()
        }
    }

    fun usuarioExistente(): Boolean {
        return _usuario.value.email.isNotEmpty()
    }

    // Devuelve el valor real de DataStore (no el initial value del StateFlow)
    suspend fun checkInitialLoginState(): Boolean = repo.readInitialLoginState()
}
