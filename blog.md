---
title: Blog
---

{% for post in site.posts %}
<section>
<h3><a href="{{ post.url }}">{{ post.title }}</a></h3>
<p>{{ post.description}}</p>
</section>
{% else %}

Upcoming posts on vba-blocks:

- Building vba-blocks
- Working with vba-blocks
- Manifest deep dive
- Lots more!

Check back soon!

{% endfor %}