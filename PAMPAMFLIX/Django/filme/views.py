from django.shortcuts import render, redirect, reverse
from .models import Filme, Usuario
from .forms import CriarContaForm, FormHomepage
from django.views.generic import TemplateView, ListView, DetailView, FormView, UpdateView
from django.contrib.auth.mixins import LoginRequiredMixin

# Create your views here.
class Homepage(FormView):
    template_name = "homepage.html"
    form_class = FormHomepage

    def get(self, request,*args, **kwargs ):
        if request.user.is_authenticated:
            return redirect('filme:homefilmes')
        else:
            return super().get(self, request,*args, **kwargs)

    def get_success_url(self):
        email = self.request.POST.get("email")
        usuarios = Usuario.objects.filter(email=email)
        if usuarios:
            return reverse('filme:login')
        else:
            return reverse('filme:criarconta')

class Homefilmes(LoginRequiredMixin, ListView):
    template_name = "homefilmes.html"
    model = Filme #objecti_list = lista de itens do modelo

class Detalhesfilmes(LoginRequiredMixin,DetailView):
    template_name = "detalhesfilme.html"
    model = Filme #object = 1 item do nosso modelo

    def get(self,request, *arg, **kwargs):
        #contabilizar uma visualização
        filme = self.get_object()
        filme.visualizacoes += 1
        filme.save()
        usuario = request.user
        usuario.filmes_vistos.add(filme)
        return super().get(self,request, *arg, **kwargs)#redireciona o usuário para a url final

    def get_context_data(self, **kwargs):
        context = super(Detalhesfilmes, self).get_context_data(**kwargs)
        #filtrar minha tabela de filmes pegando os filmes cuja categoria é igual a categoria do filme da pagina (object)
        #self.get_object = object
        filmes_relacionados = Filme.objects.filter(categoria = self.get_object().categoria)[0:5]
        context['filmes_relacionados'] = filmes_relacionados
        return context


class Pesquisafilme(LoginRequiredMixin, ListView):
    template_name = 'pesquisa.html'
    model = Filme

    def get_queryset(self):
        termo_pesquisa = self.request.GET.get('query')
        if termo_pesquisa:
            object_list = Filme.objects.filter(titulo__icontains = termo_pesquisa)
            return object_list
        else:
            return None


class Paginaperfil(LoginRequiredMixin, UpdateView):
    template_name = 'editarperfil.html'
    model = Usuario
    fields = ['first_name','last_name','email']

    def get_success_url(self):
        return reverse('filme:homefilmes')



class Criarconta( FormView):
    template_name = 'criarconta.html'
    form_class  = CriarContaForm

    def form_valid(self,form):
        form.save()
        return super().form_valid(form)

    def get_success_url(self):
        return reverse('filme:login')


class Fotos(LoginRequiredMixin,TemplateView):
    template_name = 'fotospampam.html'