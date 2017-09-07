---
title: Blog
---

{% for post in site.posts %}
<section>
<h3><a href="{{ post.url }}">{{ post.title }}</a></h3>
<p>{{ post.description}}</p>
</section>
{% endfor %}